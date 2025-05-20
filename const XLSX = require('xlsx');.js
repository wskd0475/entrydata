const XLSX = require('xlsx');
const readline = require('readline-sync');
const fs = require('fs');
const path = require('path');




// Main class for time registration
class TimeRegistrationApp {
  constructor() {
    this.entries = [];
    this.excelFile = '';
    this.csvFile = 'time_entries.csv';
  }

  // Load data from Excel file
  async loadFromExcel(filePath) {
    try {
      console.log(`Loading data from ${filePath}...`);
      const workbook = XLSX.readFile(filePath);
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      
      // Convert Excel data to JSON
      this.entries = XLSX.utils.sheet_to_json(worksheet);
      console.log(`Successfully loaded ${this.entries.length} entries.`);
      return true;
    } catch (error) {
      console.error(`Error loading Excel file: ${error.message}`);
      return false;
    }
  }

  // Save data to CSV file
  async saveToCSV() {
    try {
      if (this.entries.length === 0) {
        console.log('No entries to save.');
        return false;
      }

      // Get headers from the first entry
      const headers = Object.keys(this.entries[0]);
      
      // Create CSV content
      let csvContent = headers.join(',') + '\n';
      
      // Add each entry as a row
      for (const entry of this.entries) {
        const row = headers.map(header => {
          // Handle values that might contain commas by wrapping in quotes
          const value = entry[header] !== undefined ? entry[header] : '';
          return typeof value === 'string' && value.includes(',') 
            ? `"${value}"` 
            : value;
        }).join(',');
        csvContent += row + '\n';
      }
      
      // Write to file
      await fs.writeFile(this.csvFile, csvContent);
      console.log(`Data saved to ${this.csvFile}`);
      return true;
    } catch (error) {
      console.error(`Error saving to CSV: ${error.message}`);
      return false;
    }
  }

  // Add a new time entry
  async addNewEntry() {
    try {
      console.log('\n=== Add New Time Entry ===');
      
      // If we have existing entries, use their structure as a template
      let newEntry = {};
      
      if (this.entries.length > 0) {
        const template = this.entries[0];
        
        // Ask for each field value
        for (const field of Object.keys(template)) {
          const answer = await readline.question(`Enter ${field}: `);
          newEntry[field] = answer;
        }
      } else {
        // If no existing entries, ask for basic time registration fields
        newEntry.date = await readline.question('Enter date (YYYY-MM-DD): ');
        newEntry.startTime = await readline.question('Enter start time (HH:MM): ');
        newEntry.endTime = await readline.question('Enter end time (HH:MM): ');
        newEntry.description = await readline.question('Enter description: ');
        newEntry.project = await readline.question('Enter project: ');
      }
      
      // Add the new entry
      this.entries.push(newEntry);
      console.log('Entry added successfully!');
      
      // Save immediately
      await this.saveToCSV();
      
      return true;
    } catch (error) {
      console.error(`Error adding new entry: ${error.message}`);
      return false;
    }
  }

  // Display all entries
  displayEntries() {
    if (this.entries.length === 0) {
      console.log('No entries to display.');
      return;
    }
    
    console.log('\n=== Current Time Entries ===');
    
    // Display each entry
    this.entries.forEach((entry, index) => {
      console.log(`\nEntry #${index + 1}:`);
      for (const [key, value] of Object.entries(entry)) {
        console.log(`${key}: ${value}`);
      }
    });
  }

  // Main application loop
  async run() {
    try {
      console.log('=== Time Registration Application ===');
      
      // Ask for Excel file path
      this.excelFile = await readline.question('Enter the path to your Excel file (or press Enter to start with empty data): ');
      
      // Load Excel data if a file path was provided
      if (this.excelFile.trim()) {
        await this.loadFromExcel(this.excelFile);
      }
      
      let running = true;
      
      while (running) {
        console.log('\n=== Menu ===');
        console.log('1. Display all time entries');
        console.log('2. Add new time entry');
        console.log('3. Save to CSV');
        console.log('4. Load from Excel file');
        console.log('5. Exit');
        
        const choice = await readline.question('Enter your choice (1-5): ');
        
        switch (choice) {
          case '1':
            this.displayEntries();
            break;
          case '2':
            await this.addNewEntry();
            break;
          case '3':
            await this.saveToCSV();
            break;
          case '4':
            const filePath = await readline.question('Enter the path to your Excel file: ');
            if (filePath.trim()) {
              await this.loadFromExcel(filePath);
            }
            break;
          case '5':
            console.log('Exiting application. Goodbye!');
            running = false;
            break;
          default:
            console.log('Invalid choice. Please try again.');
        }
      }
      
      // Close readline interface
      readline.close();
    } catch (error) {
      console.error(`Application error: ${error.message}`);
      readline.close();
    }
  }
}

// Create and run the application
const app = new TimeRegistrationApp();
app.run();s