// Smoke tests — minimal sanity checks that the runtime files load and look right.
// This widget runs in the browser via CDN-loaded Vue (no bundler), so most logic
// lives inside a single createApp({ setup() {...} }) and can't be unit-tested
// without a refactor. These tests serve mainly as a CI heartbeat to catch
// syntax errors / structural regressions, and as scaffolding for future
// per-feature tests as we extract testable helpers.
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const ROOT = resolve(__dirname, '..');
const read = (p) => readFileSync(resolve(ROOT, p), 'utf-8');

describe('app.js — structural sanity', () => {
  const src = read('app.js');

  it('parses as valid JavaScript', () => {
    // `new Function(...)` throws SyntaxError on malformed code.
    // Wrap to suppress runtime execution (we just want parsing).
    expect(() => new Function(src + '\nreturn true')).not.toThrow();
  });

  it('declares the grist.ready() entry point at top-level', () => {
    expect(src).toMatch(/grist\.ready\(\s*{/);
  });

  it('configures DOMPurify with a strict tag allow-list', () => {
    expect(src).toContain('ALLOWED_TAGS');
    // ensure script-ish tags are NOT in the allow-list
    const match = src.match(/ALLOWED_TAGS:\s*\[([^\]]*)\]/);
    expect(match).not.toBeNull();
    const allowed = match[1];
    expect(allowed).not.toMatch(/'script'/);
    expect(allowed).not.toMatch(/'iframe'/);
    expect(allowed).not.toMatch(/'object'/);
  });
});

describe('index.html — structural sanity', () => {
  const src = read('index.html');

  it('loads Vue and DOMPurify from CDN', () => {
    expect(src).toMatch(/vue@3.*vue\.global\.js/);
    expect(src).toMatch(/dompurify@3/);
  });

  it('loads the Grist plugin API', () => {
    expect(src).toContain('grist-plugin-api.js');
  });

  it('declares the form container the Vue app mounts on', () => {
    expect(src).toMatch(/<div\s+id="app"/);
  });
});
