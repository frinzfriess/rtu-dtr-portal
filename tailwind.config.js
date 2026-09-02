export default {
  darkMode: 'class',
  content: ["./index.html", "./src/**/*.{js,ts,jsx,tsx}"],
  theme: { 
    extend: {
      boxShadow: {
        'glow': '0 0 25px -5px rgba(6, 182, 212, 0.25)',
        'neon': '0 0 20px -3px rgba(16, 185, 129, 0.3)',
        'soft': '0 10px 30px -5px rgba(0, 0, 0, 0.08)'
      }
    } 
  },
  plugins: [],
}