# Introduction

This project includes the files necessary to write client side Dynamics 365/Dataverse/Power Platform code in Typescript,
which will then be transpiled into Javascript to deploy to Dataverse. It also includes a number of samples of useful code.

# Usage Instructions

Javascript files generated from this project will need to be added to Dataverse, either manually or via a tool such as
Webresources Manager in XRMToolbox, however, after this point there are several scenarios supported.

NOTE: This project should be added to your Visual Studio solution as a "Webresources" named Project or similar

1. Automated build - add this project to a larger solution, where the javascript files are mapped into a solution pack during a
   pipeline build.

2. Manual build - once the files are in Dataverse, the solution can be exported then imported into downstream environments.

From a build automation, reliability and efficiency perspective, scenario 1 is the better of the 2, but if your ALM maturity
is not yet at that point, this project can still be used to write Typescript, which is superior to Javascript.
https://www.geeksforgeeks.org/difference-between-typescript-and-javascript/

Either way, the usage of this project is relatively simple. When a typescript file is compiled, it will output a .js file in the
\rh\_\js folder. This file should then be added to the relevant form, and the relevant handler added to the event handler in question
in PowerPlatform solution tooling. Please note that this compiled Javascript includes a class and namespace structure, so for example:

-   Your file: rh_account.ts, Class Account, function OnLoad
-   Your event handler to register in Power Platform: rh.Account.OnLoad
-   Always tick the "Pass execution context as first parameter" option under the parameters section.

If you wish to replace the "rh\_" prefix, then do the following:

1. Edit the rollup.config.js file and replace "_const namespace = "rh";_" with your prefix
2. Edit the filenames in the ts folder
3. Run the build task!
