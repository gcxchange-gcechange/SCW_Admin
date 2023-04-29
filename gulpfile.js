'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);
build.addSuppression(`Warning - tslint - src/webparts/scwAdmin/components/ScwAdmin.tsx(303,21): error no-shadowed-variable: Shadowed name: 'arr'`);
build.addSuppression(`Warning - tslint - src/webparts/scwAdmin/components/ScwAdmin.tsx(393,21): error no-function-expression: Use arrow function instead of function expression`);
build.addSuppression(`Warning - tslint - src/webparts/scwAdmin/components/ScwAdmin.tsx(393,555): error no-function-expression: Use arrow function instead of function expression`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.initialize(require('gulp'));
