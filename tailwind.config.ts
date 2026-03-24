import type { Config } from "tailwindcss";

const config: Config = {
  content: ["./src/**/*.{js,jsx,ts,tsx}", "./src/taskpane/index.html"],
  theme: {
    extend: {
      colors: {
        // VarSync brand palette — used for variable highlight colors
        varsync: {
          cyan: "#00C8E8",
          amber: "#E8A020",
          lime: "#A8FF00",
          red: "#FF3B3B",
          purple: "#A78BFA",
          orange: "#FB923C",
          pink: "#F472B6",
          green: "#34D399",
        },
      },
      width: {
        taskpane: "320px",
      },
    },
  },
  plugins: [],
};

export default config;
