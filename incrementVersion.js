const fs = require("fs");
const path = require("path");

// Get the package file
const packageFile = require("./package.json");

// Increment the version #
const versionInfo = packageFile.version.toString().split('.');
versionInfo[2]++;
if (versionInfo[2] > 10) {
    versionInfo[1]++;
    versionInfo[2] -= 10;
}
if (versionInfo[1] > 10) {
    versionInfo[0]++;
    versionInfo[1] -= 10;
}
const version = versionInfo.join('.');

// Log
console.log("New Version: " + version);

// Update the package file
packageFile.version = version;
fs.writeFileSync("./package.json", JSON.stringify(packageFile, null, 2));
console.log("SPFx package file updated");

// Get the SPFx package-solution file
const packageSolutionFile = require("./config/package-solution.json");
packageSolutionFile.solution.version = "0." + version;

// Update the SPFx package-solution file
fs.writeFileSync("./config/package-solution.json", JSON.stringify(packageSolutionFile, null, 2));
console.log("SPFx package-solution file updated");
