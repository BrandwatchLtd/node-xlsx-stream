"use strict";

const xlsx_stream = require("../");
const archiver = require("archiver");
const assert = require("assert");
const concat = require("concat-stream");
const fs = require("fs");
const _ = require("lodash");
const path = require("path");
const through = require("through");
const unzipper = require("unzipper");
const xlsx = require("node-xlsx");
const { describe, it, before } = require("mocha");

describe("xlsx-stream", () => {
  const dirPath = path.join(__dirname, "tmp");
  before(async () => {
    await fs.mkdir(dirPath, { recursive: true }, (err) => {
      if (err) {
        return console.error(`Error creating directory: ${err.message}`);
      }
      console.log(`Directory ${dirPath} created successfully`);
    });
  });

  describe("archiver", () => {
    it("should create a zip archive", (done) => {
      const filePath = path.join(dirPath, "test.zip");
      const zip = archiver("zip");
      zip.pipe(
        concat((data) => {
          fs.writeFileSync(filePath, data);
          fs.createReadStream(filePath)
            .pipe(unzipper.Parse())
            .on("entry", (entry) => {
              const fileName = entry.path;
              const type = entry.type; // 'Directory' or 'File'
              if (type === "File") {
                console.log(`Zip file contains ${path.basename(fileName)}`);
                assert(
                  ["0.txt", "a.txt", "b.txt"].includes(fileName),
                  `Unexpected file: ${fileName}`,
                );
              }
              entry.autodrain();
            })
            .on("close", done);
        }),
      );

      const stream = through();

      process.nextTick(() => {
        zip.append(stream, { name: "0.txt" });
        zip.append("aaa", { name: "a.txt" });
        zip.append("bbb", { name: "b.txt" });
        zip.finalize();

        process.nextTick(() => {
          stream.write("ccc");
          stream.end("ddd");
        });
      });
    });
  });
  describe("array input", () => {
    it("should generate xlsx and match number of rows", (done) => {
      const filePath = path.join(dirPath, "array.xlsx");
      const x = xlsx_stream();
      const expectedRows = [
        ["String", "てすと", "&'\";<>", "&amp;"],
        ["Integer", 1, 2, -3],
        ["Float", 1.5, 0.3, 0.123456789e23],
        ["Boolean", true, false],
        ["Date", new Date()],
        ["2 Decimals Built-in format #2", { v: 1.5, nf: "0.00" }],
        ["Time Built-in format #18", { v: 1.5, nf: "h:mm AM/PM" }],
        ["Percentage Built-in format #9", { v: 0.5, nf: "0.00%" }],
        ["Percentage Custom format", { v: 0.5, nf: "00.000%" }],
        ["Duration 36 hours format #46", { v: 1.5, t: "n", nf: "[h]:mm:ss" }],
        ["Formula", { v: "ok", f: "CONCATENATE(A1,B2)" }],
      ];

      x.on("end", () => {
        fs.access(filePath, fs.constants.F_OK, (err) => {
          assert.strictEqual(err, null, `File does not exist: ${filePath}`);
          const workSheetsFromFile = xlsx.parse(filePath);
          const rows = workSheetsFromFile[0].data;
          assert.strictEqual(
            rows.length,
            expectedRows.length,
            `Number of rows does not match. Expected ${expectedRows.length}, but got ${rows.length}`,
          );
          done();
        });
      });

      x.on("finalize", () => console.log("FINALIZE:", arguments));
      x.pipe(fs.createWriteStream(filePath));
      expectedRows.forEach((row) => x.write(row));
      x.end();
    });
  });

  describe("Multiple sheets", () => {
    it("should create multiple sheets", (done) => {
      const filePath = path.join(dirPath, "multi.xlsx");
      const x = xlsx_stream();
      x.on("end", () => {
        fs.access(filePath, fs.constants.F_OK, (err) => {
          assert.strictEqual(err, null, `File does not exist: ${filePath}`);
          const workSheetsFromFile = xlsx.parse(filePath);
          const sheetNames = workSheetsFromFile.map((sheet) => sheet.name);
          assert(
            sheetNames.includes("1st sheet"),
            'Sheet "1st sheet" does not exist',
          );
          assert(
            sheetNames.includes("２枚目のシート"),
            'Sheet "２枚目のシート" does not exist',
          );
          done();
        });
      });
      x.on("finalize", () => console.log("FINALIZE:", arguments));
      x.pipe(fs.createWriteStream(filePath));

      const sheet1 = x.sheet("1st sheet");
      sheet1.write(["This", "is", "my", "first", "worksheet"]);
      sheet1.end();

      const sheet2 = x.sheet("２枚目のシート");
      sheet2.write(["これが", "２枚目の", "ワークシート", "です"]);
      sheet2.end();

      x.finalize();
    });
  });

  describe("character set support", () => {
    const filePath = path.join(dirPath, "charactersets.xlsx");
    const expectedRows = [
      ["あああ"],
      ["080-1234-5678", "これはテストです"],
      [123, 1.24, 1e50, "←数値型"],
      ["&;\"'", new Date().toISOString().split("T")[0]], // Adjusting date format for comparison
    ];

    it("should create an xlsx file and verify the contents of the rows", (done) => {
      const x = xlsx_stream();
      const out = fs.createWriteStream(filePath);
      x.pipe(out);
      expectedRows.forEach((row) => x.write(row));
      x.end();
      out.on("finish", () => {
        const workSheetsFromFile = xlsx.parse(filePath);
        const rows = workSheetsFromFile[0].data;

        // Adjusting date format for comparison
        rows[3][1] = rows[3][1].split("T")[0];
        assert.deepStrictEqual(
          rows,
          expectedRows,
          `Rows do not match. Expected ${JSON.stringify(expectedRows)}, but got ${JSON.stringify(rows)}`,
        );
        done();
      });
    });
  });

  describe("empty string", () => {
    const filePath = path.join(dirPath, "emptyString.xlsx");
    const expectedRows = [
      ['Place Holder', ''],
    ];

    it("should create an xlsx file and preserve empty string", (done) => {
      const x = xlsx_stream();
      const out = fs.createWriteStream(filePath);
      x.pipe(out);
      expectedRows.forEach((row) => x.write(row));
      x.end();
      out.on("finish", () => {
        const workSheetsFromFile = xlsx.parse(filePath);
        const rows = workSheetsFromFile[0].data;

        assert.deepStrictEqual(
          rows,
          expectedRows,
          `Rows do not match. Expected ${JSON.stringify(expectedRows)}, but got ${JSON.stringify(rows)}`,
        );
        done();
      });
    });
  });

  describe("Large dataset", () => {
    it("should handle large dataset", (done) => {
      const randRows = 10000 + _.random(10000);
      console.log(`Generating ${randRows} rows`);
      const filePath = path.join(dirPath, "large.xlsx");
      const x = xlsx_stream();
      x.on("end", () => {
        fs.access(filePath, fs.constants.F_OK, (err) => {
          assert.strictEqual(err, null, `File does not exist: ${filePath}`);
          const workSheetsFromFile = xlsx.parse(filePath);
          const rows = workSheetsFromFile[0].data;
          assert.strictEqual(
            rows.length,
            randRows,
            `Number of rows does not match. Expected 10000, but got ${rows.length}`,
          );
          done();
        });
      });
      x.on("finalize", () => console.log("FINALIZE:", arguments));
      x.pipe(fs.createWriteStream(filePath));
      for (let i = 0; i < randRows; i++) {
        x.write([i, i * 2, i * 3, i * 4, i * 5]);
      }
      x.end();
    });
  });
});
