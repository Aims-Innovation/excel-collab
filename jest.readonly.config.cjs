module.exports = {
  testEnvironment: 'jsdom',
  testMatch: ['**/src/__tests__/ReadOnlyEntry.test.tsx'],
  transform: {
    '^.+\\.[tj]sx?$': ['ts-jest', {
      tsconfig: { module: 'CommonJS', allowJs: true, isolatedModules: true },
    }],
  },
  transformIgnorePatterns: ['node_modules/(?!uuid/)'],
  moduleNameMapper: { '\\.css$': '<rootDir>/test-support/css-mock.cjs' },
  setupFilesAfterEnv: ['@testing-library/jest-dom'],
};
