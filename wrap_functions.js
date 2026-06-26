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
    // We are looking for the function definition: function submitAbsence(formData) { ... }
    const regex = new RegExp(`function \\s+${funcName}\\s*\\([^)]*\\)\\s*{[\\s\\S]*?catch\\s*\\(([^)]+)\\)\\s*{`, 'g');

    content = content.replace(regex, (match, errVar) => {
        // If it already has notifyAdminOfError, skip
        if (match.includes('notifyAdminOfError')) return match;

        return `${match}\n    notifyAdminOfError("${funcName}", ${errVar});`;
    });
});

fs.writeFileSync('code.gs', content, 'utf8');
