import fs, { PathLike } from 'fs';
import unzip from 'unzip';
import rimraf from 'rimraf';

const getAllExcelFiles = (
    path: PathLike,
) => {
    const files = fs.readdirSync(path);
    return files.filter((file) => file.endsWith('.xlsx'));
};

const unzipAllBinFiles = async (
    excelFilesLocation: string,
    excelFiles: string[],
) => {
    if (!fs.existsSync('temp')) {
        fs.mkdirSync(`temp`);
    }
    const allUnzipJobs = excelFiles.map((excelFile) => {
        return new Promise((resolve, reject) => {
            fs.createReadStream(`${excelFilesLocation}/${excelFile}`)
            .pipe(unzip.Parse())
            .on('entry', (entry) => {
                const fileName = entry.path;
                if (fileName.startsWith('xl/embeddings/')) {
                    if (!fs.existsSync(`./temp/${excelFile}`)) {
                        fs.mkdirSync(`./temp/${excelFile}`);
                    }
                    const fileNameFormat = fileName.slice(fileName.lastIndexOf('/') + 1)
                  entry.pipe(fs.createWriteStream(`./temp/${excelFile}/${fileNameFormat}`, { flags: 'w' }));
                } else {
                  entry.autodrain();
                }
            })
            .on('close', () => resolve(true))
        });
    });
    await Promise.all(allUnzipJobs);
    return true;
};

const extractOriginalFiles = async () => {
    const folders = fs.readdirSync('temp');
    if (!fs.existsSync('out')) {
        fs.mkdirSync('out');
    }
    const allExtractJobs = folders.map((folder) => {
        const path = `temp/${folder}`;
        const files = fs.readdirSync(path);
        return files.map((file) => {
            return new Promise((resolve, reject) => {
                const binary = fs.readFileSync(`${path}/${file}`);
                const headerPosition = binary
                    .indexOf('4d5a90000300000004000000ffff0000b8000000000000004000', 0, 'hex');
                const header = binary.slice(0, headerPosition);
                const fileNameStart = header.lastIndexOf('0200', undefined, 'hex');
                const fileNameRange = header.slice(fileNameStart + 2);
                const fileNameEnd = fileNameRange.indexOf('00', 0, 'hex');
                const fileName = fileNameRange.slice(0, fileNameEnd).toString('ascii');
                console.log(fileName)
                const outPutfolder = `out/${folder}`;
                if (!fs.existsSync(outPutfolder)) {
                    fs.mkdirSync(outPutfolder);
                }
                fs.writeFile(`out/${folder}/${fileName}`, binary.slice(headerPosition), (error) => {
                    if (error) {
                        reject(error);
                    }
                    resolve();
                });
            });
        })
    })
    await Promise.all(allExtractJobs);
    return true;
};

const removeTempFiles = () => {
    return new Promise((resolve, reject) => {
        rimraf('temp', (error) => {
            if (error) {
                reject(error);
            }
            resolve();
        });
    });
};

(async () => {
    try {
        const excelFilesLocation = process.argv[2] || '.';
        console.log(`[Event]: Getting all excel files.`);
        const allExcelFiles = getAllExcelFiles(excelFilesLocation);
        if (!allExcelFiles.length) {
            throw new Error('No .xlsx file found.');
        }
        console.log(`[Event]: Getting all excel files succeed.`);
        console.log(`[Event]: Getting all object files.`);
        await unzipAllBinFiles(excelFilesLocation, allExcelFiles);
        console.log(`[Event]: Getting all object files succeed.`);
        console.log(`[Event]: Start output files.`);
        await extractOriginalFiles();
        console.log(`[Event]: Output files succeed.`);
    } catch (error) {
        console.log(`[Error]: ${error.message}`);
    } finally {
        console.log(`[Event]: Removing temp files.`);
        await removeTempFiles();
        console.log(`[Event]: Remove temp files succeed.`);
    }
})()
