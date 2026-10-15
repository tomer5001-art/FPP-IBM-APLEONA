import { defineConfig } from 'vitest/config';

// Database tests: run the SQL migrations against a real Postgres and test the
// RPC functions (permissions, calculations, audit log).
// Requires TEST_DATABASE_URL, e.g. postgres://user:pass@localhost:5432/attendance_test
export default defineConfig({
  test: {
    include: ['db-tests/**/*.test.ts'],
    environment: 'node',
    testTimeout: 30000,
    hookTimeout: 60000,
    fileParallelism: false,
  },
});
