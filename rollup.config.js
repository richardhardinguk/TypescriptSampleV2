// Note: There is a rollup-plugin-typescript that will consume TypeScript directly.  
// However, for the reason explained in https://github.com/rollup/rollup-plugin-typescript/issues/28
// enum constants are not expanded in the modules, and are left as arrays in the output 
// JavaScript.
// 
// To improve on that we use a two phase build, first from TypeScript to JavaScript
// ES2015 modules, and then bundling the ES2015 JavaScript modules using rollup.
// This is suggested here: https://github.com/rollup/rollup-plugin-typescript/issues/28#issuecomment-242999859
// and the repo with it setup is here:
// https://github.com/maxdavidson/typescript-library-boilerplate
// 
// Folder structure:
//     ts/        - TypeScript sources, any files at the top level will be bundled to modules
//     build/     - Intermediary ES2015 modules
//     ebecs_/js  - Output Directory of bundles (configured in outputDir below)

import { glob } from 'glob';
import { basename } from 'path';
import sourcemaps from 'rollup-plugin-sourcemaps';
import nodeResolve from 'rollup-plugin-node-resolve';
import nodeGlobals from 'rollup-plugin-node-globals';
import nodeBuiltins from 'rollup-plugin-node-builtins';
// import commonjs from 'rollup-plugin-commonjs';    // Uncomment if need to bundle external commonjs modules
// import { terser } from 'rollup-plugin-terser';    // Uncomment if 
// import { resolveTypeReferenceDirective } from 'typescript';

// Indicates where the bundles will be output, this is the final resting place of the webresource
// that will be loaded by Dynamics.
const outputDir = "rh_/js";

// The namespace that the bundle will be under, the iife variable
const namespace = "rh";


// Find all typescript files, we will create a "bundle" for each.
// In the simple case we'll just produce a bundle for each entity.
const bundles = glob.sync("ts/*.ts");

// The rollup configuration
let bundleList = [];

bundles.forEach((filepath) => {
    const file = basename(filepath, ".ts");
    console.log(`Adding ${file} to config.`);
    bundleList.push({
        input: `build/${file}.js`,
        output: {
            extend: true,
            file: `${outputDir}/${file}.js`,
            format: "iife",
            name: namespace,
            sourcemap: "inline"
        },
        plugins: [
            sourcemaps(), nodeResolve(), nodeGlobals(), nodeBuiltins()  //, commonjs()

            // Optionally can minimise code by adding terser() to pipeline.
            // sourcemaps(), nodeResolve(), nodeGlobals(), nodeBuiltins(), terser({
            //      toplevel: false,
            //      compress: {
            //          // Ensures that === comparisons aren't optimised away, which keeps the
            //          // PowerApps checker happy.
            //          comparisons: false,
            //          // Ensures that console.* calls are removed, which keeps the solution
            //          // checker happy.
            //          drop_console: true,
            //      },
            // })
        ]
    });
});

export default bundleList;