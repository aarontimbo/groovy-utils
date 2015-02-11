#!/usr/bin/env groovy

import org.apache.poi.hssf.usermodel.*
import org.apache.poi.ss.usermodel.*
import org.joda.time.LocalDateTime
import org.joda.time.format.DateTimeFormat
import org.joda.time.format.DateTimeFormatter


class XlsConverter {

    private static final INPUT_EXT  = '.xls'
    private static final OUTPUT_EXT = '.csv'
    DateTimeFormatter frmt = DateTimeFormat.forPattern("yyyyMMdd'T'HHmmss")

    private final String defaultDir = '/tmp'
	String inputPath
    String outputPath
    File inputFile

	XlsConverter(args) {
		parseArgs(args)
	}

    def parseArgs(args) {
        def cli = new CliBuilder(usage: "convert_xls_to_csv.groovy -f path/to/file/infile.xls -o path/to/outfile.csv")

        cli.with {
            f longOpt: 'inputFilePath', args: 1, argName: 'inputPath', "An XLS file to be converted to CSV"
            o longOpt: 'outputFilePath', args: 1, argName: 'outputPath', "Output file for CSV formatted data"
        }

        def options = cli.parse(args)

        if (options.'inputFilePath') {
        	inputPath = options.'inputFilePath'
        	assert inputPath =~ /${INPUT_EXT}$/
            inputFile = new File(inputPath)
        }

        if (!inputPath) {
            cli.usage()
            System.exit(-1)
        } else if (!inputFile.exists()) {
            println "Unable to find input file from given path: $inputFilePath"
            cli.usage()
            System.exit(-1)
        }

        inputFile = new File(inputPath)

        if (options.'outputFilePath') {
            outputPath = options.'outputFilePath'
            assert inputPath =~ /${OUTPUT_EXT}/
        } else {
            def outFileName = inputFile.name[0 .. inputFile.name.lastIndexOf('.') - 1]
            outputPath = "${defaultDir}/${outFileName}_${frmt.print(new LocalDateTime())}${OUTPUT_EXT}"    
        }



        println "Formatting ${inputPath} as CSV..."

        println "\nWriting to ${outputPath}"
    }

    def parseWorkbook() {
        Workbook wb = WorkbookFactory.create(inputFile);
        println "parsing workbook::${wb}"

        List<HSSFSheet> rateSheets = wb.sheets.findAll{ it.sheetName =~ /Rate Table/ }
        rateSheets.each { HSSFSheet sheet ->
            parseSpreadsheet(sheet)
        }
    }

    def parseSpreadsheet(HSSFSheet sheet) {
        println "\tparsing sheet::${sheet.sheetName}"
    }

}

new XlsConverter(args).parseWorkbook()
