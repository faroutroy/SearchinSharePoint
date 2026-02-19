'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppressMatch(
  new RegExp(`^Warning - ${['tslint'].join('|')}`, 'i')
);

build.initialize(require('gulp'));
