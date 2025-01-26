#!/usr/bin/env node
const { Command } = require('commander');
const { name, version } = require('../package.json');
const excelParser = require('../src/service-guide-parser');
// const excelParser = require('../src/parser');

const program = new Command();
program
  .name(name)
  .description('Tools to transform excel file to js variant file')
  .version(version)
  .option('-s --source <filepath>', 'exist excel file name')
  .option('-t --target <filepath>', 'generate js file name')
  .action((opts) => {
    // excelParser.transformExcel(`${process.cwd()}/${opts.source}`, opts.target);
    excelParser.transformXinjiangServiceGuideExcel(`${process.cwd()}/${opts.source}`, opts.target);
  });

program.parse();