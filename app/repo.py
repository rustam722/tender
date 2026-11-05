"""Репозиторий справочников: единственная точка доступа к базе.

Только отображение датаклассов в строки и обратно, без бизнес-логики.
Каждый save_* -- одна транзакция: при ошибке ничего не записывается.
Деньги и часы лежат в колонках «DECIMAL TEXT» и возвращаются как Decimal;
суммировать их нужно в Python, а не через SQL SUM().
"""
import json
import sqlite3
from collections.abc import Callable
from dataclasses import asdict, dataclass, field, fields, is_dataclass, replace
from decimal import Decimal

from app.models import (
    Asset, CustomerProfile, DocSpec, Employee, ExperienceRecord, LotItem, Organization,
    OrgSettings, Service, Tender, TenderCostItem, TenderLaborRow,
)

# Значения по умолчанию для настроек калькуляции (общие ограничения плана).
# НДС 5 % -- ставка УСН микропредприятия.
DEFAULT_SETTINGS = OrgSettings(
    insurance_rate=Decimal("0.30"),
    injury_rate=Decimal("0.002"),
    overhead_rate=Decimal("0.05"),
    profit_rate=Decimal("0.10"),
    vat_rate=Decimal("0.05"),
    total_rounding_digits=0,  # целые рубли, как ROUND(..., +D88) в КП.xlsx
)

_SETTINGS_ID = 1  # в org_settings ровно одна строка


def _dataclass_to_json(obj):
    """json.dumps(default=...): датакласс внутри JSON-поля -- объект с полями."""
    if is_dataclass(obj) and not isinstance(obj, type):
        return asdict(obj)
    raise TypeError(f"Нельзя сохранить в JSON: {type(obj).__name__}")


@dataclass(frozen=True)
class _Entity:
    """Как датакласс ложится в таблицу."""
    table: str
    cls: type
    label: str                       # для сообщения об отсутствующей записи
    json_fields: tuple[str, ...] = ()  # колонки с JSON
    # JSON-колонка -> как собрать значение поля из разобранного JSON; для полей,
    # где внутри лежат не голые строки и числа, а датаклассы
    decoders: dict[str, Callable] = field(default_factory=dict)
    extra_fields: tuple[str, ...] = ()  # поля без колонки (хранятся отдельно)

    @property
    def columns(self) -> list[str]:
        return [f.name for f in fields(self.cls)
                if f.name != "id" and f.name not in self.extra_fields]


_ORGANIZATION = _Entity("organization", Organization, "Организация")
_EMPLOYEE = _Entity("employee", Employee, "Сотрудник")
_ASSET = _Entity("asset", Asset, "Ресурс")
_EXPERIENCE = _Entity("experience", ExperienceRecord, "Опыт")
_SERVICE = _Entity("service", Service, "Услуга", extra_fields=("norms",))
_PROFILE = _Entity("customer_profile", CustomerProfile, "Профиль заказчика",
                   json_fields=("documents", "docs"),
                   decoders={"docs": lambda raw: {
                       name: DocSpec(**spec) for name, spec in raw.items()}})
_TENDER = _Entity("tender", Tender, "Тендер",
                  json_fields=("deal_answers", "experience_ids"))
_LOT_ITEM = _Entity("lot_item", LotItem, "Позиция лота")
_TENDER_LABOR = _Entity("tender_labor", TenderLaborRow, "Строка трудозатрат")
_TENDER_COST_ITEM = _Entity("tender_cost_item", TenderCostItem,
                            "Строка материалов или субподряда")
_SETTINGS = _Entity("org_settings", OrgSettings, "Настройки")


class Repo:
    def __init__(self, conn: sqlite3.Connection):
        self.conn = conn

    # ---- общие помощники -------------------------------------------------

    def _upsert(self, ent: _Entity, obj, row_id: int | None) -> int:
        """INSERT ... ON CONFLICT(id) DO UPDATE. Не коммитит."""
        cols = ent.columns
        values = []
        for name in cols:
            value = getattr(obj, name)
            if name in ent.json_fields:
                value = json.dumps(value, ensure_ascii=False,
                                   default=_dataclass_to_json)
            values.append(value)
        names = ", ".join(["id"] + cols)
        marks = ", ".join("?" * (len(cols) + 1))
        updates = ", ".join(f"{c} = excluded.{c}" for c in cols)
        cur = self.conn.execute(
            f"INSERT INTO {ent.table} ({names}) VALUES ({marks}) "
            f"ON CONFLICT(id) DO UPDATE SET {updates}",
            [row_id, *values])
        return row_id if row_id is not None else cur.lastrowid

    def _save(self, ent: _Entity, obj) -> int:
        with self.conn:
            return self._upsert(ent, obj, obj.id)

    def _build(self, ent: _Entity, row: dict):
        kwargs = {}
        for name in ent.columns:
            value = row[name]
            if name in ent.json_fields:
                value = json.loads(value)
                if name in ent.decoders:
                    value = ent.decoders[name](value)
            kwargs[name] = value
        return ent.cls(id=row["id"], **kwargs)

    def _select(self, ent: _Entity, where: str = "", params=(),
                order: str = "id") -> list:
        sql = f"SELECT * FROM {ent.table}"
        if where:
            sql += f" WHERE {where}"
        cur = self.conn.execute(f"{sql} ORDER BY {order}", params)
        names = [d[0] for d in cur.description]
        return [self._build(ent, dict(zip(names, row))) for row in cur.fetchall()]

    def _get(self, ent: _Entity, row_id: int):
        found = self._select(ent, "id = ?", (row_id,))
        if not found:
            raise LookupError(f"{ent.label}: запись с id={row_id} не найдена")
        return found[0]

    def _delete(self, ent: _Entity, row_id: int) -> None:
        # Дочерние записи убирает ON DELETE CASCADE / SET NULL в схеме.
        with self.conn:
            self.conn.execute(f"DELETE FROM {ent.table} WHERE id = ?", (row_id,))

    # ---- организации -----------------------------------------------------

    def save_organization(self, org: Organization) -> int:
        return self._save(_ORGANIZATION, org)

    def get_organization(self, id: int) -> Organization:
        return self._get(_ORGANIZATION, id)

    def list_organizations(self, role: str | None = None) -> list[Organization]:
        if role is None:
            return self._select(_ORGANIZATION)
        return self._select(_ORGANIZATION, "role = ?", (role,))

    def find_organization_by_inn(self, inn: str) -> Organization | None:
        found = self._select(_ORGANIZATION, "inn = ?", (inn,))
        return found[0] if found else None

    def delete_organization(self, id: int) -> None:
        self._delete(_ORGANIZATION, id)

    # ---- сотрудники ------------------------------------------------------

    def save_employee(self, employee: Employee) -> int:
        return self._save(_EMPLOYEE, employee)

    def get_employee(self, id: int) -> Employee:
        return self._get(_EMPLOYEE, id)

    def list_employees(self) -> list[Employee]:
        return self._select(_EMPLOYEE)

    def delete_employee(self, id: int) -> None:
        self._delete(_EMPLOYEE, id)

    # ---- материально-технические ресурсы ---------------------------------

    def save_asset(self, asset: Asset) -> int:
        return self._save(_ASSET, asset)

    def get_asset(self, id: int) -> Asset:
        return self._get(_ASSET, id)

    def list_assets(self) -> list[Asset]:
        return self._select(_ASSET)

    def delete_asset(self, id: int) -> None:
        self._delete(_ASSET, id)

    # ---- опыт ------------------------------------------------------------

    def save_experience(self, record: ExperienceRecord) -> int:
        return self._save(_EXPERIENCE, record)

    def get_experience(self, id: int) -> ExperienceRecord:
        return self._get(_EXPERIENCE, id)

    def list_experience(self, service_id: int | None = None) -> list[ExperienceRecord]:
        if service_id is None:
            return self._select(_EXPERIENCE)
        return self._select(_EXPERIENCE, "service_id = ?", (service_id,))

    def delete_experience(self, id: int) -> None:
        self._delete(_EXPERIENCE, id)

    # ---- услуги и нормы --------------------------------------------------

    def save_service(self, service: Service) -> int:
        with self.conn:
            sid = self._upsert(_SERVICE, service, service.id)
            # Нормы пересоздаются целиком: убранная должность не должна остаться.
            self.conn.execute("DELETE FROM service_norm WHERE service_id = ?", (sid,))
            self.conn.executemany(
                "INSERT INTO service_norm (service_id, position, hours) VALUES (?, ?, ?)",
                [(sid, position, hours) for position, hours in service.norms.items()])
        return sid

    def _with_norms(self, services: list[Service]) -> list[Service]:
        for service in services:
            rows = self.conn.execute(
                "SELECT position, hours FROM service_norm "
                "WHERE service_id = ? ORDER BY rowid", (service.id,))
            service.norms = {row[0]: row[1] for row in rows}
        return services

    def get_service(self, id: int) -> Service:
        return self._with_norms([self._get(_SERVICE, id)])[0]

    def list_services(self) -> list[Service]:
        return self._with_norms(self._select(_SERVICE))

    def find_service_by_code(self, code: str) -> Service | None:
        found = self._with_norms(self._select(_SERVICE, "code = ?", (code,)))
        return found[0] if found else None

    def delete_service(self, id: int) -> None:
        self._delete(_SERVICE, id)

    # ---- профили заказчиков ----------------------------------------------

    def save_customer_profile(self, profile: CustomerProfile) -> int:
        return self._save(_PROFILE, profile)

    def get_customer_profile(self, id: int) -> CustomerProfile:
        return self._get(_PROFILE, id)

    def find_customer_profile(self, customer_id: int) -> CustomerProfile | None:
        found = self._select(_PROFILE, "customer_id = ?", (customer_id,))
        return found[0] if found else None

    def list_customer_profiles(self) -> list[CustomerProfile]:
        return self._select(_PROFILE)

    def delete_customer_profile(self, id: int) -> None:
        self._delete(_PROFILE, id)

    # ---- тендеры и позиции лота ------------------------------------------

    def save_tender(self, tender: Tender) -> int:
        return self._save(_TENDER, tender)

    def get_tender(self, id: int) -> Tender:
        return self._get(_TENDER, id)

    def list_tenders(self) -> list[Tender]:
        return self._select(_TENDER)

    def delete_tender(self, id: int) -> None:
        self._delete(_TENDER, id)

    def save_lot_item(self, item: LotItem) -> int:
        return self._save(_LOT_ITEM, item)

    def get_lot_item(self, id: int) -> LotItem:
        return self._get(_LOT_ITEM, id)

    def list_lot_items(self, tender_id: int) -> list[LotItem]:
        return self._select(_LOT_ITEM, "tender_id = ?", (tender_id,),
                            order="position, id")

    def delete_lot_item(self, id: int) -> None:
        self._delete(_LOT_ITEM, id)

    # ---- трудозатраты тендера --------------------------------------------

    def save_tender_labor(self, row: TenderLaborRow) -> int:
        return self._save(_TENDER_LABOR, row)

    def list_tender_labor(self, tender_id: int) -> list[TenderLaborRow]:
        return self._select(_TENDER_LABOR, "tender_id = ?", (tender_id,),
                            order="position, id")

    def delete_tender_labor(self, id: int) -> None:
        self._delete(_TENDER_LABOR, id)

    # ---- материалы и субподряд тендера -----------------------------------

    def save_tender_cost_item(self, item: TenderCostItem) -> int:
        return self._save(_TENDER_COST_ITEM, item)

    def list_tender_cost_items(self, tender_id: int,
                               kind: str | None = None) -> list[TenderCostItem]:
        if kind is None:
            return self._select(_TENDER_COST_ITEM, "tender_id = ?", (tender_id,),
                                order="position, id")
        return self._select(_TENDER_COST_ITEM, "tender_id = ? AND kind = ?",
                            (tender_id, kind), order="position, id")

    def delete_tender_cost_item(self, id: int) -> None:
        self._delete(_TENDER_COST_ITEM, id)

    # ---- настройки калькуляции -------------------------------------------

    def get_settings(self) -> OrgSettings:
        cur = self.conn.execute("SELECT * FROM org_settings WHERE id = ?",
                                (_SETTINGS_ID,))
        row = cur.fetchone()
        if row is None:
            return replace(DEFAULT_SETTINGS)  # копия: общий объект не отдаём
        names = [d[0] for d in cur.description]
        data = dict(zip(names, row))
        return OrgSettings(**{c: data[c] for c in _SETTINGS.columns})

    def save_settings(self, settings: OrgSettings) -> None:
        with self.conn:
            self._upsert(_SETTINGS, settings, _SETTINGS_ID)
