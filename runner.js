const fs = require('fs');
let code = fs.readFileSync('code.gs', 'utf8');
code += `
module.exports = {
  getStaffRosterForAdmin,
  saveStaffMemberAdmin,
  deleteStaffMemberAdmin,
  bulkUpsertStaffRoster
};
`;
fs.writeFileSync('temp_code.js', code);
