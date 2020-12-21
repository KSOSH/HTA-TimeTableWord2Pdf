module.exports = function(grunt) {
	require('load-grunt-tasks')(grunt);
	require('time-grunt')(grunt);
	const path = require('path');
	var pkg = grunt.file.readJSON('package.json');
	grunt.initConfig({
		globalConfig: {},
		pkg: pkg,
		less: {
			main: {
				options : {
					compress: true,
					ieCompat: true
				},
				files: {
					'css/main.css': [
						'css/main.less'
					]
				}
			}
		},
		htmlImagesDataUri: {
			dist: {
				src: ['test/*.html'],
				dest: 'test/',
				options: {
					target: ['images/*.*', 'help/*.*'],
					baseDir: './'
				}
			}
		},
		pug: {
			files: {
				options: {
					pretty: '',//'\t',
					separator:  '',//'\n'
				},
				files: {
					"test/index.html": ['index.pug'],
				}
			}
		},
		exec: {
			hta: {
				cmd: 'copy /y /B "' + path.resolve(__dirname + '/icon.ico') + '" + "' + path.resolve(__dirname + '/test/index.html') + '" "' + path.resolve(__dirname + '/TimeTable2pdf.hta') + '"'
			}
		}
	});
	grunt.registerTask('default',["less", "pug", "htmlImagesDataUri", "exec"]);
}