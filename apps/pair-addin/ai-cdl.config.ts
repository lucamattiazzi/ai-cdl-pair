import { defineConfig } from "@ai-cdl/cli/config";

export default defineConfig({
  app: {
    id: "4dc1f4b5-e74d-4b9a-87d1-d36d78a2bd4f",
    name: "AI-CDL Pair",
    description: "Connect your own agent to Excel with bounded reads and approval before writes",
    version: "0.1.1",
    providerName: "AI-CDL",
  },
  office: {
    hosts: ["Workbook"],
    permissions: "ReadWriteDocument",
    requirements: { ExcelApi: "1.13" },
  },
  taskpane: {
    developmentUrl: "https://localhost:3000",
    productionUrl: process.env.PAIR_PUBLIC_ORIGIN ?? "https://pair.example.com",
    developmentCommand: "pnpm dev",
    supportPath: "/support.html",
  },
  commands: {
    label: "Open AI-CDL Pair",
    groupLabel: "AI-CDL Pair",
    icon: "./public/icon.png",
    icons: {
      16: "./public/icon-16.png",
      32: "./public/icon-32.png",
      64: "./public/icon-64.png",
      80: "./public/icon-80.png",
    },
  },
});
