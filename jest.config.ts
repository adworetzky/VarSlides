import type { Config } from "jest";

const config: Config = {
  preset: "ts-jest",
  testEnvironment: "jsdom",
  roots: ["<rootDir>/tests"],
  transform: {
    "^.+\\.tsx?$": [
      "ts-jest",
      {
        tsconfig: "tsconfig.jest.json",
      },
    ],
  },
  moduleNameMapper: {
    "^@/(.*)$": "<rootDir>/src/$1",
    // stub CSS/asset imports in tests
    "\\.(css|less|scss|sass)$": "<rootDir>/tests/__mocks__/styleMock.js",
  },
  setupFiles: ["<rootDir>/tests/setupOffice.ts"],
};

export default config;
