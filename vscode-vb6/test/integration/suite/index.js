'use strict';
const path = require('path');
const Mocha = require('mocha');

exports.run = function run() {
    const mocha = new Mocha({ ui: 'bdd', color: false, timeout: 60000 });
    mocha.addFile(path.join(__dirname, 'project.test.js'));
    return new Promise((resolve, reject) => {
        mocha.run((failures) => (failures ? reject(new Error(`${failures} integration test(s) failed`)) : resolve()));
    });
};
