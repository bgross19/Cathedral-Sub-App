const fs = require('fs');
let code = fs.readFileSync('Index.html', 'utf8');

code = code.replace(
  `              } else {
                alert("Database Error: " + (result ? result.error : "Unknown issue"));
                btn.disabled = false;
                btn.innerHTML = 'Save Changes';
              }`,
  `              } else {
                alert(result ? result.error : "Database Error");
                btn.disabled = false;
                btn.innerHTML = 'Save Changes';
              }`
);

code = code.replace(
  `              } else {
                alert("Database Error: " + (result ? result.error : "Unknown issue"));
                btn.disabled = false;
                btn.innerHTML = 'Submit Request';
              }`,
  `              } else {
                alert(result ? result.error : "Database Error");
                btn.disabled = false;
                btn.innerHTML = 'Submit Request';
              }`
);

fs.writeFileSync('Index.html', code);
