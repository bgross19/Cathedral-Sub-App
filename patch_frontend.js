const fs = require('fs');
let code = fs.readFileSync('Index.html', 'utf8');

code = code.replace(
  `              if (result && result.error && (result.error.includes('filled by someone else') || result.error.includes('server is currently busy'))) {
                  loadTodaysOpenJobs();
              }
            }
          })
          .withFailureHandler(err => {
            alert("Error: " + err.message);
            btn.disabled = false;
            btn.innerText = originalText;
          })
          .assignSubToPeriod(absenceId, period, subName);`,
  `              if (result && result.error && (result.error.includes('filled by someone else') || result.error.includes('server is currently busy'))) {
                  loadQuickCover();
                  loadAdminDataRawSilent();
              }
            }
          })
          .withFailureHandler(err => {
            alert("Error: " + err.message);
            btn.disabled = false;
            btn.innerText = originalText;
          })
          .assignSubToPeriod(absenceId, period, subName);`
);

fs.writeFileSync('Index.html', code);
