module.exports = function(grunt) {
  grunt.initConfig({
    pkg: grunt.file.readJSON('package.json'),
    site: grunt.file.readYAML('_config.yml'),

    assemble: {
      options: {
        flatten: true,
        assets: '<%= site.dest %>',

        pkg: '<%= pkg %>',
        site: '<%= site %>',
        data: ['content/**/*.json'],
        content: ['content/**/*.md'],

        partials: '<%= site.includes %>/*.hbs',
        layoutdir: '<%= site.layouts %>',
        layoutext: '<%= site.layoutext %>',
        layout: '<%= site.layout %>',

        compose: {cwd: 'content'},

        marked: {
          process: true,
          prefix: 'lang-'
        }
      },
      docs: {
        expand: true,
        cwd: '<%= site.pages %>/',
        src: ['**/*.hbs'],
        dest: '<%= site.dest %>/'
      }
    },

    clean: {
      docs: ['<%= site.dest %>']
    },

    connect: {
      options: {
        port: 9000,
        livereload: 35729,
        hostname: 'localhost'
      },
      livereload: {
        options: {
          open: true,
          base: ['<%= site.dest %>']
        }
      }
    },

    copy: {
      assets: {
        files: [
          {
            expand: true,
            cwd: '<%= site.assets %>',
            src: ['**'],
            dest: '<%= assemble.options.assets %>/'  
          }
        ]
      },
      scripts: {
        files: [
          {
            expand: true,
            cwd: '<%= site.scripts %>',
            src: ['**'],
            dest: '<%= assemble.options.assets %>/js/'
          }
        ]
      }
    },

    jshint: {
      options: {
        jshintrc: '.jshintrc'
      },
      all: ['Gruntfile.js']
    },

    less: {
      docs: {
        src: ['styles/main.less'],
        dest: '<%= assemble.options.assets %>/css/main.css'
      }
    },

    watch: {
      options: {
        livereload: true
      },
      styles: {
        files: ['<%= site.styles %>/**/*.less'],
        tasks: ['less']
      },
      content: {
        files: ['<%= site.content %>/**/*.md'],
        tasks: ['assemble']
      },
      templates: {
        files: ['<%= site.templates %>/**/*.hbs'],
        tasks: ['assemble']
      },
      assets: {
        files: ['<%= site.assets %>/**/*'],
        tasks: ['copy:assets']
      },
      scripts: {
        files: ['<%= site.scripts %>/**/*'],
        tasks: ['copy:scripts']
      },
      pages: {
        files: ['<%= site.pages %>/**/*.hbs'],
        tasks: ['assemble']
      }
    },

    'gh-pages': {
      options: {
        base: '_gh-pages'
      },
      src: ['**']
    },

    prettify: {
      docs: {
        files: [
          {
            expand: true,
            cwd: '<%= site.dest %>',
            src: '*.html',
            dest: '<%= site.dest %>',
            ext: '.html'
          }
        ]
      }
    }
  });

  // Load plugins
  grunt.loadNpmTasks('assemble');
  grunt.loadNpmTasks('grunt-contrib-clean');
  grunt.loadNpmTasks('grunt-contrib-copy');
  grunt.loadNpmTasks('grunt-contrib-connect');
  grunt.loadNpmTasks('grunt-contrib-jshint');
  grunt.loadNpmTasks('grunt-contrib-less');
  grunt.loadNpmTasks('grunt-contrib-watch');
  grunt.loadNpmTasks('grunt-gh-pages');
  grunt.loadNpmTasks('grunt-prettify');

  // Register tasks
  grunt.registerTask('publish', ['default', 'gh-pages']);

  grunt.registerTask('design', [
    'clean',
    'copy',
    'less',
    'assemble',
    'connect',
    'watch'
  ]);

  grunt.registerTask('default', [
    'jshint',
    'clean',
    'copy',
    'less',
    'assemble'
  ]);
};
