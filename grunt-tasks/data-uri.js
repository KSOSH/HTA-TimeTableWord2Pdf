module.exports = function(grunt){
	var fs = require('fs'),
		path = require('path'),
		datauri = require('datauri'),
		regex = /(?:(?:["']?data-uri\()(.*?)(?:\)["']?))/,
		util = grunt.util,
		gruntFileDIr = path.resolve('./'),
		expandFiles = grunt.file.expandFiles ? grunt.file.expandFiles : function(files) {
			return grunt.file.expand({filter: 'isFile'}, files);
		};
	grunt.registerMultiTask('data-uri', 'Convert your file to base64', function() {
		var options = this.options(),
			srcFiles = expandFiles(this.data.src),
			destDir = path.resolve(this.data.dest),
			haystak = [],
			done = this.async();

		async function srcFilesLoop(arr, content){
			var arss = [],
				baseDir = options.baseDir ? path.resolve(options.baseDir) : path.resolve(path.dirname(src));
			for (let i=0; i<arr.length; i++){
				let path_url = path.join(baseDir, arr[i]),
					url = arr[i].replace(/\\/, '/'),
					data = await datauri(path_url),
					pattern = '(?:(?:["\']?data-uri\\()' + url.replace('/', '\\/') + '(?:\\)["\']?))',
					reg = new RegExp(pattern, 'g');
				content = content.replace(reg, '"' + data + '"');
			}
			return content;
		}

		async function readFiles(arrFiles){
			for (let i=0; i < arrFiles.length; i++){
				var src = arrFiles[i],
					content = String(fs.readFileSync(src)),
					matches = content.match(new RegExp(regex.source, 'g')),
					baseDir = options.baseDir ? path.resolve(options.baseDir) : path.resolve(path.dirname(src)),
					outputTo = destDir + '/' + path.basename(src),
					uris, skrab;
				if(!matches){
					grunt.file.write(outputTo, content);
					continue;
				}
				uris = util._.uniq(matches.map(function(m){
					var r = m.match(regex)[1];
					return r;
				}));
				content = await srcFilesLoop(uris, content);
				grunt.file.write(outputTo, content);
			}
		}

		readFiles(srcFiles).then(done).catch(function(e){
			grunt.log.error(e);
			done();
		});
	});

}