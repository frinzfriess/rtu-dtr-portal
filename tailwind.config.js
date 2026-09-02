export default {
  content: ["./index.html", "./src/**/*.{js,ts,jsx,tsx}"],
  theme: { 
    extend: {
      boxShadow: {
        'soft': '0 4px 20px -2px rgba(0,0,0,0.05)',
        'float': '0 10px 30px -5px rgba(79,70,229,0.15)',
      }
    } 
  },
  plugins: [],
}