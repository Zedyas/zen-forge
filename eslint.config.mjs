import eslint from '@eslint/js'
import tseslint from 'typescript-eslint'

export default tseslint.config(
  {
    ignores: ['node_modules/**', 'out/**', 'plans/**'],
  },
  eslint.configs.recommended,
  ...tseslint.configs.recommended,
  {
    files: ['src/renderer/src/**/*.{ts,tsx}'],
    rules: {
      'no-restricted-imports': [
        'error',
        {
          patterns: [
            {
              group: ['node:*'],
              message: 'Renderer modules must access files through FileService.',
            },
          ],
        },
      ],
    },
  },
  {
    files: ['src/renderer/src/**/*.{ts,tsx}'],
    ignores: ['src/renderer/src/services/file/**'],
    rules: {
      'no-restricted-syntax': [
        'error',
        {
          selector: "MemberExpression[property.name='desktop']",
          message: 'Only the FileService adapter may access the preload bridge.',
        },
      ],
    },
  },
)
