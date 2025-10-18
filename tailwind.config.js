/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{ts,tsx}'],
  theme: {
    extend: {
      colors: {
        brand: '#53d3a4',
        'brand-dark': '#3abf90',
        navy: '#0f1b25',
        'navy-card': '#172635'
      },
      boxShadow: {
        raised: '0 6px 14px rgba(0,0,0,.25), inset 0 1px 0 rgba(255,255,255,.06)',
        'raised-sm': '0 4px 10px rgba(0,0,0,.18), inset 0 1px 0 rgba(255,255,255,.05)'
      }
    },
  },
  plugins: [],
}