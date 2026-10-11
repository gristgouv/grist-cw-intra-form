// Structural tests for the custom-select component (search + chips + a11y),
// lazy-load perf for large ref tables, and ref-dropdown filtering.
// Same approach as smoke.test.js: assert presence of identifying tokens
// rather than running the Vue app.
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const ROOT = resolve(__dirname, '..');
const read = (p) => readFileSync(resolve(ROOT, p), 'utf-8');

describe('custom-select — a11y baked into the template', () => {
  const html = read('index.html');
  const src = read('app.js');

  it('uses a <button> trigger (combobox lives on the search input)', () => {
    // Sole combobox is the search input — avoids two comboboxes (APG).
    const trigger = html.match(
      /<button\s+type="button"\s+class="custom-select-display"[\s\S]*?>/
    )?.[0];
    expect(trigger).toBeTruthy();
    expect(trigger).not.toContain('role="combobox"');
  });

  it('puts role="combobox" on the search input (open state)', () => {
    expect(html).toMatch(/class="search-input"[\s\S]*?role="combobox"/);
    expect(html).toMatch(/aria-autocomplete="list"/);
  });

  it('hides the trigger from the a11y tree while the dropdown is open', () => {
    expect(html).toMatch(/:aria-hidden="openDropdown === element\.fieldName \? 'true' : null"/);
  });

  it('renders an aria-labelled <ul role="listbox">', () => {
    expect(html).toMatch(/role="listbox"/);
  });

  it('declares aria-multiselectable on the listbox for multi-selects', () => {
    expect(html).toContain('aria-multiselectable');
  });

  it('uses .sr-only spans for multi-select hints (count + selected labels)', () => {
    expect(html).toMatch(/class="sr-only"/);
    expect(html).toContain("selected_' + element.fieldName");
    expect(src).toMatch(/function getSelectedValuesText/);
  });

  it('exposes searchDescribedBy for the open search combobox', () => {
    expect(src).toMatch(/function searchDescribedBy/);
    expect(html).toContain(':aria-describedby="searchDescribedBy(element)"');
  });

  it('leaves aria-activedescendant unset on initial open (activeIndex = null)', () => {
    // So SRs announce the search combobox before any option.
    expect(src).toMatch(/activeIndex\[fieldName\]\s*=\s*null/);
  });
});

describe('lazy-load perf for large ref tables', () => {
  const src = read('app.js');

  it('defines ensureRefDataLoaded for first-open lazy fetch', () => {
    expect(src).toMatch(/function ensureRefDataLoaded/);
  });

  it('caps displayed options via MAX_DISPLAYED_OPTIONS', () => {
    expect(src).toMatch(/MAX_DISPLAYED_OPTIONS\s*=\s*\d+/);
  });
});

describe('ref-dropdown filter', () => {
  const src = read('app.js');
  const html = read('index.html');

  it('declares refDropdownFilterPopup state', () => {
    expect(src).toContain('refDropdownFilterPopup');
  });

  it('renders the 🔗 (chain) icon in the config modal', () => {
    expect(html).toContain('&#128279;');
  });
});
