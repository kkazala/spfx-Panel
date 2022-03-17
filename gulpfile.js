'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};
// ********* disable tslint *******
// https://github.com/pnp/pnpjs/issues/1636
build.tslintCmd.enabled = false;
// ********* disable tslint *******

build.initialize(require('gulp'));
