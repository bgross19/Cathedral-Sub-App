const fs = require('fs');

let content = fs.readFileSync('code.gs', 'utf8');

const funcs = [
    "updateSettings",
    "editUserRole",
    "deleteUserRole",
    "addUserRole",
    "assignSubToPeriod",
    "cancelMySubDuty",
    "updateAbsence",
    "submitAbsence",
    "cancelAbsence"
];

funcs.forEach(funcName => {
    // A slightly less strict regex that looks for catch blocks inside the function
    const regex = new RegExp(`(function\\s+${funcName}\\s*\\([\\s\\S]*?catch\\s*\\([a-zA-Z0-9_]+\\)\\s*{)`, 'g');
    let matches = content.match(regex);
    if(matches) {
        console.log(`Found catch block in ${funcName}`);
    } else {
        console.log(`NO catch block found in ${funcName}`);
    }
});
