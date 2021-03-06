'use strict';
var path = require('path');
var assert = require('yeoman-assert');
var helpers = require('yeoman-test');

describe('generator-sharepoint-webpart-extension:app', function () {
  before(function () {
    return helpers.run(path.join(__dirname, '../generators/app'))
      .withPrompts({
        name: 'foo',
        description: 'bar',
        authorName: 'FooBar',
        keywords: ['']
      })
      .toPromise();
  });

  it('creates files', function () {
    assert.file([
      '.gitignore',
      'gulpfile.js',
      'package.json',
      'package.nuspec',
      'tsconfig.json',
      'typings.json',
      'webpack.config.js',
      'NuGet_lib/Readme.txt',
      'NuGet_tools/install.ps1',
      'src/index.html',
      'src/foo.dwp',
      'src/assets/LogoLarge.png',
      'src/assets/LogoSmall.png',
      'src/js/AppPartPropertyUIOverride.ts',
      'src/js/index.tsx',
      'src/js/components/Counter.tsx'
    ]);
  });
});
