const { defineConfig } = require('eslint-define-config');

module.exports = defineConfig({
  root: true,
  env: {
    browser: true,
    node: true,
    es6: true,
  },
  parserOptions: {
    parser: '@typescript-eslint/parser',
    ecmaVersion: 2020,
    sourceType: 'module',
  },
  extends: [
    'plugin:@typescript-eslint/recommended',
    'eslint:recommended',
    'airbnb-base',
  ],
  plugins: [
    '@typescript-eslint',
  ],
  rules: {
    'new-cap': 'off',
    'no-undef': 'off',
    'no-shadow': 'off',
    'no-bitwise': 'off',
    'max-len': [2, 200],
    'no-console': 'off',
    'no-debugger': 'off',
    'import/extensions': 'off',
    'no-fallthrough': 'off',
    'no-param-reassign': [
      'error',
      {
        props: true,
        ignorePropertyModificationsFor: [
          'opts', 'props', 'output', 'style', 'paragraph', 'content', 'run', 'cell',
          'row', 'table', 'elem', 'target',
        ],
      },
    ],
    'default-param-last': 'off',
    'import/no-unresolved': 'off',
    'no-use-before-define': 'off',
    '@typescript-eslint/ban-types': 'off',
    'max-classes-per-file': [2, 5],
    '@typescript-eslint/no-shadow': 'off',
    '@typescript-eslint/ban-ts-comment': 'off',
    'no-underscore-dangle': ['error', {
      allow: ['_options', '_parser', '_zip', '_text',
        '_xml', '_xmlDocument', '_documentParser'],
    }],
    '@typescript-eslint/no-unused-vars': ['error', { argsIgnorePattern: '^_' }],
    'no-unused-vars': ['error', { argsIgnorePattern: '^_' }],
  },
  overrides: [
    {
      files: ['*.js', 'tsup.config.ts'],
      rules: {
        'global-require': 'off',
        '@typescript-eslint/no-var-requires': 'off',
        'import/no-extraneous-dependencies': 'off',
      },
    },
    {
      files: ['src/renderer.ts', 'src/utils.ts'],
      rules: {
        'guard-for-in': 'off',
        'operator-assignment': 'off',
        'block-scoped-var': 'off',
        'no-unused-expressions': 'off',
        'no-param-reassign': 'off',
        'no-dupe-else-if': 'off',
        'no-prototype-builtins': 'off',
        'no-mixed-operators': 'off',
        'no-restricted-syntax': 'off',
        '@typescript-eslint/no-explicit-any': 'off',
      },
    },
    {
      files: ['src/types.ts', 'src/document/types.ts', 'src/document/section.ts'],
      rules: {
        'no-unused-vars': 'off',
      },
    },
  ],
});
