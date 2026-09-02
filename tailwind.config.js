export default {
  content: ["./index.html", "./src/**/*.{js,ts,jsx,tsx}"],
  theme: { 
    extend: {
      colors: {
        cyber: {
          950: '#030712',
          900: '#0f172a',
          800: '#1e293b',
        }
      },
      boxShadow: {
        'glow': '0 0 25px -5px rgba(6, 182, 212, 0.25)',
        'neon': '0 0 20px -3px rgba(16, 185, 129, 0.3)'
      }
    } 
  },
  plugins: [],
}