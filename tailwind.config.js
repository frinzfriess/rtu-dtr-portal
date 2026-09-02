export default {
  darkMode: 'class',
  content: ["./index.html", "./src/**/*.{js,ts,jsx,tsx}"],
  theme: { 
    extend: {
      boxShadow: {
        'glow': '0 0 25px -5px rgba(6, 182, 212, 0.25)',
        'soft': '0 8px 30px rgb(0,0,0,0.04)'
      }
    } 
  },
  plugins: [],
}