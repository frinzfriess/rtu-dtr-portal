export default {
  content: ["./index.html", "./src/**/*.{js,ts,jsx,tsx}"],
  theme: { 
    extend: {
      colors: {
        dark: {
          950: '#090d16',
          900: '#111827',
          800: '#1f2937',
          700: '#374151',
          primary: '#3b82f6', // Professional blue
          accent: '#10b981',  // Emerald
        }
      },
      boxShadow: {
        'glow': '0 0 25px -5px rgba(59, 130, 246, 0.2)',
        'soft': '0 10px 30px -5px rgba(0, 0, 0, 0.3)'
      }
    } 
  },
  plugins: [],
}