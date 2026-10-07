// Keys pressed while an input method (注音, 倉頡, 速成 …) is still composing a word pick or
// confirm its characters: Enter chooses the character, Esc drops the candidates. They are never
// the panel's own Enter (search, replace, apply) or Esc (close, cancel) — persona-300 A-7. The
// test: `isComposing`, and keyCode 229 for browsers that send the confirming Enter with it only
// (Safari, older Chrome).

/** Whether a key belongs to an input method still composing text. */
export function composing(e: KeyboardEvent): boolean {
  return e.isComposing || e.keyCode === 229;
}

/** A keydown handler that ignores keys of a composing input method. */
export function outsideIme<E extends KeyboardEvent>(handler: (e: E) => void): (e: E) => void {
  return (e) => {
    if (!composing(e)) handler(e);
  };
}

/** Enter (not while composing): prevents the default and runs `action`. */
export function onEnter(action: () => void): (e: KeyboardEvent) => void {
  return (e) => {
    if (e.key !== 'Enter' || composing(e)) return;
    e.preventDefault();
    action();
  };
}
