'use strict';
var yeoman = require('yeoman-generator');
var chalk = require('chalk');
var yosay = require('yosay');
var path = require('path');
var _ = require('lodash');

function makeGeneratorName(name) {
  name = _.kebabCase(name);
  return name;
}

module.exports = yeoman.Base.extend({
  prompting: function () {
    // Have Yeoman greet the user.
    this.log(yosay(
      'Welcome to the ' + chalk.red('SharePoint WebPart Extension') + ' generator!'
    ));

    var prompts = [{
      name: 'name',
      message: 'Your SharePoint Web Part extension package name',
      default: makeGeneratorName(path.basename(process.cwd())),
      validate: function (str) {
        return str.length > 0;
      }
    },{
      name: 'classname',
      message: 'Your SharePoint Web Part extension JavaScript class name',
      default: path.basename(process.cwd()),
      validate: function (str) {
        return str.length > 0;
      }
    }, {
      name: 'displayname',
      message: 'Your SharePoint Web Part extension display name',
      default: path.basename(process.cwd()),
      validate: function (str) {
        return str.length > 0;
      }
    }, {
        name: 'description',
        message: 'Description'
      }, {
        name: 'authorName',
        message: 'Author\'s Name',
        default: this.user.git.name(),
        store: true
      }, {
        name: 'keywords',
        message: 'Package keywords (comma to split)',
        filter: function (words) {
          return words.split(/\s*,\s*/g);
        }
      }];

    return this.prompt(prompts).then(function (props) {
      // To access props later use this.props.someAnswer;
      this.props = props;
    }.bind(this));
  },

  writing: function () {
	// Root Folder
    this.fs.copy(this.templatePath('gitignore'), this.destinationPath('.gitignore'));
    this.fs.copyTpl(this.templatePath('bower.json'), this.destinationPath('bower.json'), this.props);
	this.fs.copy(this.templatePath('gulpfile.template'), this.destinationPath('gulpfile.js'));
    this.fs.copyTpl(this.templatePath('package.json'), this.destinationPath('package.json'), this.props);
    this.fs.copyTpl(this.templatePath('package.nuspec'), this.destinationPath('package.nuspec'), this.props);
	this.fs.copy(this.templatePath('tsconfig.json'), this.destinationPath('tsconfig.json'));
	this.fs.copy(this.templatePath('typings.json'), this.destinationPath('typings.json'));

	// Folder NuGet_lib
	this.fs.copy(this.templatePath('NuGet_lib/Readme.txt'), this.destinationPath('NuGet_lib/Readme.txt'));
	
	// Folder NuGet_tools
	this.fs.copy(this.templatePath('NuGet_tools/install.ps1'), this.destinationPath('NuGet_tools/install.ps1'));
	
	// Folder src
	this.fs.copyTpl(this.templatePath('src/index.html'), this.destinationPath('src/index.html'), this.props);
	this.fs.copyTpl(this.templatePath('src/template.dwp'), this.destinationPath('src/' + this.props.name + '.dwp'), this.props);
	
	// Folder src/assets
	this.fs.copy(this.templatePath('src/assets/LogoLarge.png'), this.destinationPath('src/assets/LogoLarge.png'));
	this.fs.copy(this.templatePath('src/assets/LogoSmall.png'), this.destinationPath('src/assets/LogoSmall.png'));
	
	// Folder src/js
	this.fs.copy(this.templatePath('src/js/AppPartPropertyUIOverride.ts'), this.destinationPath('src/js/AppPartPropertyUIOverride.ts'));
	this.fs.copy(this.templatePath('src/js/extension.d.ts'), this.destinationPath('src/js/extension.d.ts'));
	this.fs.copyTpl(this.templatePath('src/js/extension.ts'), this.destinationPath('src/js/extension.ts'), this.props);
  },
  
  // writing_package: function () {
  //   var pkg = this.fs.readJSON(this.destinationPath('package.json'), {});
  //   extend(pkg, {
  //     dependencies: {
  //       'yeoman-generator': '^0.23.0',
  //       chalk: '^1.0.0',
  //       yosay: '^1.0.0'
  //     },
  //     devDependencies: {
  //       'yeoman-test': '^1.0.0',
  //       'yeoman-assert': '^2.0.0'
  //     }
  //   });
  //   pkg.keywords = pkg.keywords || [];
  //   pkg.keywords.push('yeoman-generator');
  //
  //   this.fs.writeJSON(this.destinationPath('package.json'), pkg);
  // },

  install: function () {
    this.installDependencies({ npm: true, bower: true });
	this.spawnCommand('typings', ['install']);
  }
});
