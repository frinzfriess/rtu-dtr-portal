export default {
  darkMode: 'class',
  content: ["./index.html", "./src/**/*.{js,ts,jsx,tsx}"],
  theme: {
    extend: {
      boxShadow: {
        'glow': '0 0 25px -5px rgba(6, 182, 212, 0.25)',
        'neon': '0 0 20px -3px rgba(16, 185, 129, 0.3)',
        'soft': '0 10px 30px -5px rgba(0, 0, 0, 0.08)',
        'inner-soft': 'inset 0 2px 4px 0 rgba(0, 0, 0, 0.06)'
      },
      keyframes: {
        fadeIn: {
          '0%': { opacity: '0', transform: 'translateY(8px)' },
          '100%': { opacity: '1', transform: 'translateY(0)' }
        },
        popIn: {
          '0%': { opacity: '0', transform: 'scale(0.96)' },
          '100%': { opacity: '1', transform: 'scale(1)' }
        }
      },
      animation: {
        fadeIn: 'fadeIn 0.4s ease-out both',
        popIn: 'popIn 0.2s ease-out both'
      }
    }
  },
  plugins: [],
}