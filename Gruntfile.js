module.exports = function(grunt) {
	grunt.loadTasks('grunt-tasks');
	require('load-grunt-tasks')(grunt);
	require('time-grunt')(grunt);
	const path = require('path');
	var pkg = grunt.file.readJSON('package.json');
	grunt.initConfig({
		globalConfig: {},
		pkg: pkg,
		"less": {
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
		"data-uri": {
			dist: {
				src: ['test/*.html'],
				dest: 'test/',
				options: {
					baseDir: __dirname
				}
			}
		},
		"pug": {
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
		"exec": {
			hta: {
				cmd: 'copy /y /B "' + path.join(__dirname, 'icon.ico') + '" + "' + path.join(__dirname, 'test/index.html') + '" "' + path.resolve(__dirname + '/TimeTable2pdf.hta') + '"'
			},
			run: {
				cmd: 'cmd /c start TimeTable2pdf.hta'
			}
		}
	});
	grunt.registerTask('default',["less", "pug", "data-uri", "exec"]);
}