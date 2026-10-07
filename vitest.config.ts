import {fileURLToPath, URL} from "node:url";
import {defineConfig} from "vitest/config";
import vue from "@vitejs/plugin-vue";

// Tests for the DOCX editor (src/papyrus). Opt-in suites over real documents:
//   DOCX_SAMPLES=<folder> npm test
export default defineConfig({
    // Components (.vue) are mounted by some tests.
    plugins: [vue()],
    resolve: {
        alias: {
            "@": fileURLToPath(new URL("./src", import.meta.url)),
        },
    },
    test: {
        environment: "jsdom",
        include: ["tests/**/*.test.ts"],
    },
});
