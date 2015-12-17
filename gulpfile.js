var gulp = require('gulp');
var shell = require('gulp-shell');

gulp.task('gen', shell.task("node src/js/gen-timesheet.js"))