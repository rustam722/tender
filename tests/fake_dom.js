// Минимальная заглушка DOM для тестов на Node: ровно столько, сколько нужно коду
// экранов из app/web/static/app.js, чтобы он ИСПОЛНЯЛСЯ (построение таблиц,
// отрисовка расчёта и результата сборки, переходы между экранами). Это не
// браузер: вид, вёрстка, фокус, прокрутка и настоящее поведение элементов она не
// проверяет, а только то, что код отрисовки доходит до конца и не падает на
// ответах сервера (с итогами, без итогов, с предупреждением).
'use strict';

class FakeNode {
  constructor() { this.parentNode = null; }
  get isConnected() {
    let node = this;
    while (node.parentNode) node = node.parentNode;
    return node.isRoot === true;
  }
  remove() {
    if (this.parentNode) this.parentNode.children = this.parentNode.children.filter((c) => c !== this);
    this.parentNode = null;
  }
}

class FakeText extends FakeNode {
  constructor(data) { super(); this.data = String(data); }
  get textContent() { return this.data; }
}

class FakeElement extends FakeNode {
  constructor(tag) {
    super();
    this.tagName = tag.toUpperCase();
    this.children = [];
    this.attributes = {};
    this.listeners = {};
    this.className = '';
    this.hidden = false;
    this.disabled = false;
    this.checked = false;
    this.open = false;
    this.returnValue = '';
    this._value = undefined;
  }

  // ---- дерево
  append(...nodes) {
    for (const node of nodes) {
      const child = node instanceof FakeNode ? node : new FakeText(node);
      if (child.parentNode) child.remove();
      child.parentNode = this;
      this.children.push(child);
    }
  }
  replaceChildren(...nodes) {
    for (const child of this.children) child.parentNode = null;
    this.children = [];
    this.append(...nodes);
  }
  get firstChild() { return this.children[0] || null; }
  get textContent() { return this.children.map((c) => c.textContent).join(''); }
  set textContent(text) { this.replaceChildren(String(text)); }

  // ---- атрибуты
  setAttribute(name, value) { this.attributes[name] = String(value); }
  getAttribute(name) { return name in this.attributes ? this.attributes[name] : null; }
  removeAttribute(name) { delete this.attributes[name]; }
  get classList() {
    const el = this;
    const names = () => el.className.split(/\s+/).filter(Boolean);
    return {
      add(...n) { el.className = [...new Set([...names(), ...n])].join(' '); },
      remove(...n) { el.className = names().filter((x) => !n.includes(x)).join(' '); },
      contains(n) { return names().includes(n); },
      toggle(n, force) {
        const on = force === undefined ? !names().includes(n) : Boolean(force);
        if (on) this.add(n); else this.remove(n);
        return on;
      },
    };
  }

  // ---- значение поля: у select -- по списку option, как в браузере
  get value() {
    if (this.tagName === 'SELECT') {
      const options = this.children.filter((c) => c.tagName === 'OPTION');
      if (this._value === undefined) return options.length ? options[0].value : '';
      const found = options.find((o) => o.value === this._value);
      return found ? found.value : '';
    }
    return this._value === undefined ? (this.attributes.value ?? '') : this._value;
  }
  set value(v) { this._value = String(v); }
  get selectedOptions() {
    return this.children.filter((c) => c.tagName === 'OPTION' && c.value === this.value);
  }

  // ---- события
  addEventListener(type, fn) { (this.listeners[type] = this.listeners[type] || []).push(fn); }
  dispatchEvent(event) {
    let node = this;
    while (node) {
      for (const fn of node.listeners ? (node.listeners[event.type] || []) : []) {
        fn.call(node, { type: event.type, target: this, currentTarget: node, preventDefault() {} });
      }
      if (!event.bubbles) break;
      node = node.parentNode;
    }
    return true;
  }
  click() { this.dispatchEvent({ type: 'click', bubbles: true }); }
  focus() {}
  scrollIntoView() {}
  showModal() { this.open = true; }
  close() { this.open = false; this.dispatchEvent({ type: 'close' }); }

  // ---- поиск
  querySelector(selector) {
    const match = /^#([\w-]+)$/.exec(selector);
    if (!match) throw new Error(`заглушка DOM умеет только #id, а запрошено ${selector}`);
    const walk = (node) => {
      for (const child of node.children || []) {
        if (child.attributes && child.attributes.id === match[1]) return child;
        const deeper = child.children && walk(child);
        if (deeper) return deeper;
      }
      return null;
    };
    return walk(this);
  }
  all(predicate, out = []) {
    for (const child of this.children) {
      if (child instanceof FakeElement) {
        if (predicate(child)) out.push(child);
        child.all(predicate, out);
      }
    }
    return out;
  }
}

class FakeEvent {
  constructor(type, init = {}) { this.type = type; this.bubbles = Boolean(init.bubbles); }
}

function createWorld({ fetch }) {
  const root = new FakeElement('html');
  root.isRoot = true;
  const body = new FakeElement('body');
  root.append(body);
  const document = {
    createElement: (tag) => new FakeElement(tag),
    createTextNode: (text) => new FakeText(text),
    getElementById: (id) => root.querySelector(`#${id}`),
    querySelector: (selector) => root.querySelector(selector),
  };
  const make = (tag, id, text) => {
    const el = new FakeElement(tag);
    if (id) el.setAttribute('id', id);
    if (text) el.textContent = text;
    body.append(el);
    return el;
  };
  make('a', 'nav-refs').setAttribute('href', '#/refs');
  make('a', 'nav-tender').setAttribute('href', '#/tender');
  make('a', 'nav-build').setAttribute('href', '#/build');
  make('div', 'current-tender');
  const screen = make('main', 'screen');
  const dialog = make('dialog', 'dialog');
  make('h2', 'dialog-title');
  make('p', 'dialog-message');
  make('ul', 'dialog-items');
  make('button', 'dialog-ok');
  make('button', 'dialog-cancel');

  const windowListeners = {};
  const store = new Map();
  let hash = '';
  const win = {
    addEventListener: (type, fn) => { (windowListeners[type] = windowListeners[type] || []).push(fn); },
    confirm: () => true,
  };
  const location = {
    get hash() { return hash; },
    set hash(value) {
      const next = value.startsWith('#') ? value : `#${value}`;
      if (next === hash) return;
      hash = next;
      setTimeout(() => (windowListeners.hashchange || []).forEach((fn) => fn({})), 0);
    },
  };
  const localStorage = {
    getItem: (k) => (store.has(k) ? store.get(k) : null),
    setItem: (k, v) => { store.set(k, String(v)); },
    removeItem: (k) => { store.delete(k); },
  };
  return { document, window: win, location, localStorage, Node: FakeNode, Event: FakeEvent, fetch,
    screen, dialog, root, store };
}

module.exports = { createWorld, FakeElement, FakeText };
