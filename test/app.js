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
      'bower.json',
      'package.json',
      'package.nuspec'
    ]);
  });
});
