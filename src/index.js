const fs = require("fs");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const { startOfMonth } = require("date-fns/startOfMonth");
const { endOfMonth } = require("date-fns/endOfMonth");
const { format } = require("date-fns/format");
const yargs = require("yargs");
const { addMonths } = require("date-fns");

// Read the docx file as binary
function loadFile(filepath) {
  return fs.readFileSync(filepath, "binary");
}

async function replaceValuesInDocx(templatePath, outputPath, values) {
  const content = loadFile(templatePath);

  const zip = new PizZip(content);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  // Replace the placeholders with the provided values
  doc.setData(values);

  try {
    // Render the document
    doc.render();
  } catch (error) {
    console.error("Error rendering document:", error);
    throw error;
  }

  const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
  });

  // Write the buffer to a file
  fs.writeFileSync(outputPath, buf);
}

function computeArguments() {
  // Configure yargs to handle command-line arguments
  const argv = yargs
    .option("template", {
      alias: "t",
      description: "Path to the template file",
      type: "string",
      demandOption: true,
    })
    .option("output", {
      alias: "o",
      description: "Path to the output file",
      type: "string",
      demandOption: true,
    })
    .option("monthOffset", {
      alias: "mO",
      description: "Month modifier 1, 2, -1, -2, etc",
      type: "number",
      default: 0,
      demandOption: true,
    })
    .option("values", {
      alias: "v",
      description: "Values to replace in the format key1=value1,key2=value2",
      type: "string",
      demandOption: true,
    })
    .help()
    .alias("help", "h").argv;

  const FORMAT = "MM/dd/yyyy";

  const TODAY = new Date();

  // Parse values from the command-line argument
  let values = argv.values.split(",").reduce(
    (acc, pair) => {
      const [key, value] = pair.split("=");
      acc[key] = value;
      return acc;
    },
    {
      dateOfIssue: format(startOfMonth(TODAY), FORMAT),
      dueDate: format(endOfMonth(TODAY), FORMAT),
    }
  );

  values = { ...values, monthOffset: argv.monthOffset };

  if (values.monthOffset !== 0) {
    const adjustedDate = addMonths(TODAY, values.monthOffset);

    values = {
      ...values,
      dateOfIssue: format(startOfMonth(adjustedDate), FORMAT),
      dueDate: format(endOfMonth(adjustedDate), FORMAT),
    };
  }

  return { argv, values };
}

function main() {
  const { argv, values } = computeArguments();

  // Call the function with the provided arguments
  replaceValuesInDocx(argv.template, argv.output, values)
    .then(() => {
      console.log("DOCX created successfully");
    })
    .catch((err) => {
      console.error("Error creating DOCX:", err);
    });
}

main();
