const fs = require('fs');
let html = fs.readFileSync('Index.html', 'utf8');

// The modal wrappers look like:
// <div id="addRoleModal" class="modal opacity-0 pointer-events-none fixed w-full h-full top-0 left-0 flex items-center justify-center z-[110]">
// <div id="detailsModal" class="modal opacity-0 pointer-events-none fixed w-full h-full top-0 left-0 flex items-center justify-center z-[110]">
// <div id="modal" class="modal opacity-0 pointer-events-none fixed w-full h-full top-0 left-0 flex items-center justify-center z-[110]">

html = html.replace(
  '<div id="addRoleModal" class="modal opacity-0 pointer-events-none fixed w-full h-full top-0 left-0 flex items-center justify-center z-[110]">',
  '<div id="addRoleModal" class="modal opacity-0 pointer-events-none fixed w-full h-full top-0 left-0 flex items-center justify-center z-[110]" onclick="if(event.target === this) attemptCloseAllModals()">'
);

html = html.replace(
  '<div id="detailsModal" class="modal opacity-0 pointer-events-none fixed w-full h-full top-0 left-0 flex items-center justify-center z-[110]">',
  '<div id="detailsModal" class="modal opacity-0 pointer-events-none fixed w-full h-full top-0 left-0 flex items-center justify-center z-[110]" onclick="if(event.target === this) attemptCloseAllModals()">'
);

html = html.replace(
  '<div id="modal" class="modal opacity-0 pointer-events-none fixed w-full h-full top-0 left-0 flex items-center justify-center z-[110]">',
  '<div id="modal" class="modal opacity-0 pointer-events-none fixed w-full h-full top-0 left-0 flex items-center justify-center z-[110]" onclick="if(event.target === this) attemptCloseAllModals()">'
);

fs.writeFileSync('Index.html', html);
