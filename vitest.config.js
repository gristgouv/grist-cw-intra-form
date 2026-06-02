// Vitest config — used by `npm test` and in CI.
// happy-dom is a lightweight browser environment (faster than jsdom) needed
// for tests that touch DOM (e.g. DOMPurify, document APIs). Pure-function
// tests don't need it but the cost is negligible.
export default {
  test: {
    environment: 'happy-dom',
    include: ['tests/**/*.test.js'],
  },
};
