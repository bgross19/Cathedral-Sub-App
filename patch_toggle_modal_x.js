const fs = require('fs');

let html = fs.readFileSync('Index.html', 'utf8');

// The "X" button for the Request Sub modal:
// <div onclick="toggleModal()" class="cursor-pointer z-50 p-2 hover:bg-gray-100 rounded-full">
// We want to change just this one, not the main request button.

html = html.replace('<div onclick="toggleModal()" class="cursor-pointer z-50 p-2 hover:bg-gray-100 rounded-full">', '<div onclick="attemptCloseAllModals()" class="cursor-pointer z-50 p-2 hover:bg-gray-100 rounded-full">');

fs.writeFileSync('Index.html', html);
