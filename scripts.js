let employees = JSON.parse(localStorage.getItem('employees')) || [];

// Function to add an employee
function addEmployee() {
  const name = document.getElementById('name').value;
  const contact = document.getElementById('contact').value;
  const nic = document.getElementById('nic').value;

  if (name && contact && nic) {
    const employee = {
      name: name,
      contact: contact,
      nic: nic,
      company: "Tecgate"
    };

    employees.push(employee); // Add the new employee to the array

    // Sort employees alphabetically by name
    employees.sort((a, b) => a.name.localeCompare(b.name));

    saveToLocalStorage(); // Save the updated array to localStorage
    renderTable(); // Update the table
    document.getElementById('employeeForm').reset(); // Clear the form
  } else {
    alert("Please fill all fields!");
  }
}

// Function to render the table
function renderTable() {
  const tbody = document.querySelector('#employeeTable tbody');
  tbody.innerHTML = ''; // Clear the table

  employees.forEach((employee, index) => {
    const row = document.createElement('tr');
    row.innerHTML = `
      <td><input type="checkbox" class="select-employee" data-index="${index}"></td>
      <td>${employee.name}</td>
      <td>${employee.contact}</td>
      <td>${employee.nic}</td>
      <td>${employee.company}</td>
      <td><button onclick="removeEmployee(${index})">Remove</button></td>
    `;
    tbody.appendChild(row); // Add the row to the table
  });
}

// Function to remove an employee
function removeEmployee(index) {
  employees.splice(index, 1); // Remove the employee from the array
  saveToLocalStorage(); // Save the updated array to localStorage
  renderTable(); // Update the table
}

// Function to save data to localStorage
function saveToLocalStorage() {
  localStorage.setItem('employees', JSON.stringify(employees));
}

// Function to generate the Excel file
function generateExcel() {
  // Get all selected employees
  const selectedEmployees = [];
  document.querySelectorAll('.select-employee:checked').forEach(checkbox => {
    const index = checkbox.getAttribute('data-index');
    selectedEmployees.push(employees[index]);
  });

  if (selectedEmployees.length === 0) {
    alert("Please select at least one employee!");
    return;
  }

  // Define the headers with proper capitalization
  const headers = [
    { header: "Name", key: "name" },
    { header: "Contact Number", key: "contact" },
    { header: "NIC", key: "nic" },
    { header: "Company", key: "company" }
  ];

  // Map the selected employee data to match the headers
  const data = selectedEmployees.map(employee => ({
    Name: employee.name,
    "Contact Number": employee.contact,
    NIC: employee.nic,
    Company: employee.company
  }));

  // Create a worksheet
  const worksheet = XLSX.utils.json_to_sheet(data, { header: headers.map(h => h.header) });

  // Create a new workbook
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Team Members");

  // Generate file and trigger download
  XLSX.writeFile(workbook, 'Team members list - Upg.xlsx');
}

// Render the table when the page loads
document.addEventListener('DOMContentLoaded', renderTable);