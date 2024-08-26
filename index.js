import newArr from './array.json' assert {
    type: 'json'
};
import newLongArr from './transformed_data.json' assert {
    type: 'json'
};

import {
    parse,
    format,
    isValid
} from 'date-fns';

function preprocessDateTime(dateTimeStr) {
    // Normalize date-time formats:
    // 1. Replace multiple spaces with a single space.
    // 2. Convert dot-separated dates to slash-separated dates.
    // 3. Ensure consistent time format (24-hour and 12-hour with AM/PM).

    // Replace multiple spaces with a single space
    let normalizedStr = dateTimeStr.replace(/\s{2,}/g, ' ');

    // Convert dot-separated dates to slash-separated dates
    normalizedStr = normalizedStr.replace(/(\d{1,2})\.(\d{1,2})\.(\d{4})/, '$1/$2/$3');

    // Ensure consistent time format (12-hour with AM/PM)
    if (!normalizedStr.match(/AM|PM/)) {
        // Assume it's a 24-hour format and adjust to 12-hour if needed
        normalizedStr = normalizedStr.replace(/(\d{1,2}):(\d{2}):(\d{2})/, '$1:$2:$3');
    }

    return normalizedStr;
}

function formatDateTime(dateTimeStr) {
    const normalizedStr = preprocessDateTime(dateTimeStr);
    console.log(dateTimeStr)
    let date;

    // Define possible date formats
    const formats = [
        "M/d/yy h:mm:ss a", // 8/15/24 1:58:00 PM
        "M/d/yy H:mm:ss", // 8/12/24 13:50:00
        "d/MM/yyyy h:mm:ss a", // 15/08/2032 8:25:00 AM
        "d/MM/yyyy HH:mm:ss", // 12/08/2024 09:07:00
        "M/d/yy HH:mm:ss" // 8/11/24 21:22:00
    ];

    // Attempt to parse the date string with different formats
    for (const formatStr of formats) {
        try {
            date = parse(normalizedStr, formatStr, new Date());
            if (isValid(date)) break; // Exit loop if a valid date is found
        } catch (e) {
            // Continue to the next format if the current one fails
        }
    }

    // Check if a valid date was found
    if (!isValid(date)) {
        throw new Error('Invalid date-time string');
    }

    // Extract day, month, hours, and minutes
    const day = format(date, 'd');
    const month = format(date, 'M'); // month is 1-based in this format
    const hours = format(date, 'HH'); // 24-hour format
    const minutes = format(date, 'mm');

    // Format the output
    return `${day}.${Number(month) >= 10 ? month : `0${month}`}.2024 ${hours}:${minutes}`;
}

function toGenitiveCase(name) {
    // Очищення від зайвих пробілів і перетворення на нижній регістр
    name = name.trim().toLowerCase();

    // Якщо ім'я закінчується на голосну
    if (/[аяиія]/.test(name.slice(-1))) {
        if (/а$/.test(name)) {
            return name.slice(0, -1) + "и";
        } else if (/ія$/.test(name)) {
            return name.slice(0, -2) + "ії";
        } else if (/я$/.test(name)) {
            return name.slice(0, -1) + "ї";
        }  else if (/й$/.test(name)) {
            return name.slice(0, -1) + "я";
        } else {
            return name + "и";
        }
    } 
    
    // Для імен на -ій
    else if (/ій$/.test(name)) {
        return name.slice(0, -2) + "ія";
    }
    
    // Для імен, що закінчуються на "ь"
    else if (/ь$/.test(name)) {
        return name.slice(0, -1) + "я";
    }
    
    // Якщо ім'я закінчується на приголосну
    else {
        return name + "а";
    }
}


const resultArr = newLongArr.map((obj) => {
    const lastName = obj["Прізвище"].charAt(0).toUpperCase() + obj["Прізвище"].slice(1).toLowerCase();
    const firstName = obj["Імя "].charAt(0).toUpperCase() + toGenitiveCase(obj["Імя "].slice(1).toLowerCase());
    const services = 'Залізничний квиток для ' + lastName + ' ' + firstName;
    const arrivalTime = formatDateTime(`${obj["Дата прибуття"]} ${obj["Час прибуття"]}`);
    const departureTime = formatDateTime(`${obj["Дата Відправлення "]} ${obj["Час відправлення"]}`);
    // const comment = `Потяг №${obj["Номер потягу"]} ${obj["Станція відправлення "]} - ${obj["Станція прибуття"]}`; // OLD COMMENT
    const comment = `Потяг №${obj["Номер потягу"]} ${obj["Станція відправлення "]} - ${obj["Станція прибуття"]}, Час відправлення: ${departureTime}, Час прибуття: ${arrivalTime}`; // NEW COMMENT

    return {
        services,
        prevCost: obj['Вартість квитка'],
        membersCount: 1,
        daysCount: 1,
        totalCost: obj['Вартість квитка'],
        comment,
        // arrivalTime: formatDateTime(`${obj["Дата прибуття"]} ${obj["Час прибуття"]}`),
        // departureTime: formatDateTime(`${obj["Дата Відправлення "]} ${obj["Час відправлення"]}`)
    }
});

import XLSX from 'xlsx';
import fs from 'fs';

// const filePath = 'download.xls';

// // Read the Excel file
// const workbook = XLSX.readFile(filePath);

// // Get the first sheet name
// const sheetName = workbook.SheetNames[1];

// // Get the first worksheet
// const worksheet = workbook.Sheets[sheetName];

// const outputFilePath = 'transformed_data.json';

// // Write the JSON string to the file

// // Convert the worksheet to JSON
// const jsonData = XLSX.utils.sheet_to_json(worksheet);

// const keyMapping = {
//     "__EMPTY": "Прізвище",
//     "__EMPTY_1": "Імя ",
//     "__EMPTY_2": "Станція відправлення ",
//     "__EMPTY_3": "Дата Відправлення ",
//     "__EMPTY_4": "Час відправлення",
//     "__EMPTY_5": "Станція прибуття",
//     "__EMPTY_6": "Дата прибуття",
//     "__EMPTY_7": "Час прибуття",
//     "__EMPTY_8": "Номер потягу",
//     "__EMPTY_9": "Вартість квитка",
//     // Add any other mappings as necessary
// };

// // Function to rename keys in an object
// function renameKeys(obj, mapping) {
//     const newObj = {};

//     console.log(obj);
//     for (const key in obj) {
//         if (obj.hasOwnProperty(key)) {
//             const newKey = mapping[key] || key; // Default to original key if not found in mapping
//             newObj[newKey] = `${obj[key]}`;
//         }
//     }
//     return newObj;
// }

// // Transform the array of objects
// const transformedData = jsonData.map(obj => renameKeys(obj, keyMapping));

// // Print the transformed data
// const jsonString = JSON.stringify(transformedData, null, 2);

// fs.writeFileSync(outputFilePath, jsonString, 'utf8');

// Define the output file path
const filePath = 'shablon.xlsx';

if (!fs.existsSync(filePath)) {
    console.error('File does not exist!');
    process.exit(1);
}

// Read the existing workbook
const workbook = XLSX.readFile(filePath);

// Get the first sheet (modify if needed)
const sheetName = workbook.SheetNames[0];
let worksheet = workbook.Sheets[sheetName];

// Convert new data to a sheet
const newWorksheet = XLSX.utils.json_to_sheet(resultArr, {
    header: Object.keys(resultArr[0])
});

if (worksheet) {
    // Get existing data
    const existingData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1
    });

    // Add headers if they don't exist
    const headers = Object.keys(resultArr[0]);
    if (existingData.length === 0) {
        existingData.push(headers);
    }

    // Append new data (excluding header from newWorksheet)
    const newData = XLSX.utils.sheet_to_json(newWorksheet, {
        header: 1
    });
    if (newData.length > 0) {
        existingData.push(...newData.slice(1)); // Append new data, excluding the header
    }

    // Convert updated data back to sheet
    worksheet = XLSX.utils.aoa_to_sheet(existingData);
} else {
    // If the sheet doesn't exist, use the new data as the initial sheet
    worksheet = newWorksheet;
}

// Replace the existing sheet with updated data
workbook.Sheets[sheetName] = worksheet;

// Write the updated workbook back to the file
XLSX.writeFile(workbook, filePath);

console.log('Data has been appended to the Excel file successfully!');


// OLD CODE
// import XLSX from 'xlsx';
// import fs from 'fs';

// // Define your data array

// // Create a new workbook and add the data to it
// const workbook = XLSX.utils.book_new();
// const worksheet = XLSX.utils.json_to_sheet(resultArr);

// // Append the worksheet to the workbook
// XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

// // Write the workbook to a file
// XLSX.writeFile(workbook, 'output.xlsx');

// console.log('Excel file has been created successfully!');