module.exports = function(grunt) {
  // Project configuration.
  grunt.initConfig({
    pkg: grunt.file.readJSON('package.json'),
    clean: {
      all: ['lib', 'tmp']
    },
    coffee: {
      all: {
        expand: true,
        cwd: 'src',
        src: ['**/*.coffee'],
        dest: 'lib',
        ext: '.js'
      }
    },
    mkdir: {
      all: {
        options: {
          create: ['tmp']
        }
      }
    },
    shell: {
      vows: {
        command: 'vows test/test.coffee --spec'
      }
    },
    watch: {
      scripts: {
        files: ['src/**/*.coffee'],
        tasks: ['coffee'],
        options: {
          spawn: false
        }
      }
    }
  });

  // Load the plugins.
  grunt.loadNpmTasks('grunt-contrib-clean');
  grunt.loadNpmTasks('grunt-contrib-coffee');
  grunt.loadNpmTasks('grunt-mkdir');
  grunt.loadNpmTasks('grunt-shell');
  grunt.loadNpmTasks('grunt-contrib-watch');

  // Default task(s).
  grunt.registerTask('default', ['clean', 'coffee', 'mkdir', 'shell:vows']);
};
