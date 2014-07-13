module.exports = function(grunt) {
  grunt.initConfig({
    'gh-pages': {
      options: {
        base: 'dist'
      },
      src: ['**']
    }
  });

  // Load plugins
  grunt.loadNpmTasks('grunt-gh-pages');

  // Register tasks
  grunt.registerTask('publish', ['gh-pages']);
};
