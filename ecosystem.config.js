module.exports = {
  apps: [
    {
      name: "client",
      script: "npm",
      args: "start",
      cwd: "./", // Adjust if your client is in a subfolder
      env: {
        NODE_ENV: "development",
      },
    },
    {
      name: "server",
      script: "npm",
      args: "run dev",
      cwd: "./Backend", // Path to your backend folder
      env: {
        NODE_ENV: "development",
      },
    },
  ],
};