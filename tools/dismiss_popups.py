from __future__ import annotations

DISMISS_JS = r"""() => {
  const reBtn = /^(принять(\s+все)?|согласен|разрешить|allow all|allow essential cookies?|accept( all)?|agree|i agree|got it|ok|ок|закрыть|close|понятно|reject all|отклонить(\s+все)?|do not consent|consent)$/i;
  const reLoose = /принять все|allow all|accept all|allow essential|отклонить все/i;
  const clicked = [];
  const tryClick = (el) => {
    if (!el) return false;
    const t = ((el.innerText || el.value || el.getAttribute("aria-label") || "") + "").trim();
    if (!t) return false;
    if (reBtn.test(t) || reLoose.test(t)) {
      try { el.click(); clicked.push(t.slice(0, 80)); return true; } catch (e) { return false; }
    }
    return false;
  };
  const walk = (root) => {
    if (!root || !root.querySelectorAll) return;
    for (const el of root.querySelectorAll("button, [role=button], a, input[type=button], input[type=submit]")) {
      tryClick(el);
    }
    const close = root.querySelector("[aria-label*=close i], [aria-label*=закрыть i], .close, [class*=cookie] button");
    if (close) tryClick(close);
  };
  walk(document);
  for (const node of document.querySelectorAll("*")) {
    if (node.shadowRoot) walk(node.shadowRoot);
  }
  for (const fr of document.querySelectorAll("iframe")) {
    try { if (fr.contentDocument) walk(fr.contentDocument); } catch (e) {}
  }
  return { clicked, remaining: document.body && /cookie|куки/i.test(document.body.innerText) };
}
"""


def dismiss_script() -> str:
    return DISMISS_JS.strip()
