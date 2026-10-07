import {fileURLToPath, URL} from "node:url";
import {defineConfig, type Plugin} from "vite";
import vue from "@vitejs/plugin-vue";

const src = (path: string) => fileURLToPath(new URL(path, import.meta.url));
const entry = src("./src/papyrus/index.ts");

/**
 * The standalone files carry their styles: the stylesheet is not written as a file of its own,
 * the script adds it to the page when it loads.
 */
const stylesInScript = (): Plugin => ({
    name: "papyrus-styles-in-script",
    apply: "build",
    enforce: "post",
    generateBundle(_, bundle) {
        let css = "";
        for (const [name, file] of Object.entries(bundle)) {
            if (file.type === "asset" && name.endsWith(".css")) {
                css += file.source;
                delete bundle[name];
            }
        }
        const add =
            `(function(){if(typeof document==="undefined")return;var s=document.createElement("style");` +
            `s.setAttribute("data-papyrus-docx","");s.textContent=${JSON.stringify(css)};document.head.appendChild(s)})();\n`;
        for (const file of Object.values(bundle)) {
            if (file.type === "chunk" && file.isEntry) file.code = add + file.code;
        }
    },
});

// npm run dev         the demo site: index.html (demo/main.ts) and playground.html (demo/playground.ts)
// npm run build       the package in dist/ (see package.json "exports"):
//                       papyrus-docx.js + style.css       for applications with a bundler
//                       papyrus-docx.standalone.js        one file with everything, as a module
//                       papyrus-docx.standalone.iife.js   one file with everything, for <script src>
//                       types/                            TypeScript declarations (vue-tsc)
// npm run build:demo  the demo site as static files in dist-demo/
export default defineConfig(({mode}) => {
    const shared = {
        plugins: [vue()],
        resolve: {
            alias: {
                "@": src("./src"),
            },
        },
    };
    if (mode === "demo") {
        // Relative, so the built demo works from any folder of a static host.
        return {
            ...shared,
            base: "./",
            build: {
                outDir: "dist-demo",
                target: "esnext",
                rollupOptions: {input: {index: src("./index.html"), playground: src("./playground.html")}},
            },
        };
    }
    if (mode === "standalone") {
        return {
            ...shared,
            plugins: [vue(), stylesInScript()],
            publicDir: false,
            // A library build leaves this to the application's bundler; here there is none.
            define: {"process.env.NODE_ENV": JSON.stringify("production")},
            build: {
                target: "esnext",
                emptyOutDir: false,
                lib: {
                    entry,
                    name: "PapyrusDocx",
                    formats: ["es", "iife"],
                    fileName: (format) => `papyrus-docx.standalone${format === "es" ? "" : ".iife"}.js`,
                },
            },
        };
    }
    return {
        ...shared,
        // The sample documents in public/ belong to the demo, not to the package.
        publicDir: mode === "production" ? false : "public",
        build: {
            target: "esnext",
            lib: {
                entry,
                formats: ["es"],
                fileName: "papyrus-docx",
            },
            rollupOptions: {
                // Left to the application: one copy of Vue and of each ProseMirror module.
                external: ["vue", "jszip", /^prosemirror-/],
            },
        },
    };
});
