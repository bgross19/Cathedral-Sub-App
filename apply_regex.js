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
    // Find catch (e) { ... } inside the function
    const regex = new RegExp(`(function\\s+${funcName}[\\s\\S]*?catch\\s*\\(([a-zA-Z0-9_]+)\\)\\s*{)`, 'g');
    content = content.replace(regex, (match, prefix, errVar) => {
        if (match.includes('notifyAdminOfError')) return match;
        return `${prefix}\n    notifyAdminOfError("${funcName}", ${errVar});`;
    });
});

fs.writeFileSync('code.gs', content, 'utf8');
