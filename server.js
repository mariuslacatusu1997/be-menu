const express = require('express');
const cors = require('cors');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 5000;


// Configure CORS
const corsOptions = {
    origin: 'http://localhost:3000', // Replace this with the URL of your React app
    methods: 'GET,POST,PUT,DELETE,FETCH',
    allowedHeaders: 'Content-Type',
};

app.use(cors(corsOptions)); // Use the CORS middleware
app.use(express.json());

const categoriesFilePath = path.join(__dirname, '../backend/src/categories.xlsx');
const productsFilePath = path.join(__dirname, '../backend/src/products.xlsx');

// Function to create a new Excel file if it doesn't exist
const createExcelFileIfNotExists = (filePath, defaultData) => {
    if (!fs.existsSync(filePath)) {
        const worksheet = XLSX.utils.json_to_sheet(defaultData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
        XLSX.writeFile(workbook, filePath);
    }
};

// Endpoint to initialize the files
app.get('/initialize', (req, res) => {
    createExcelFileIfNotExists(categoriesFilePath, [{ id: 1, name: 'Default Category 1' }]);
    createExcelFileIfNotExists(productsFilePath, []);
    res.send('Files initialized');
});

// Endpoint to get categories
app.get('/categories', (req, res) => {
    const workbook = XLSX.readFile(categoriesFilePath);
    const sheetName = workbook.SheetNames[0];
    const categoriesData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
    res.json(categoriesData);
});

// Endpoint to add a new category
app.post('/categories', (req, res) => {
    const newCategory = req.body;
    const workbook = XLSX.readFile(categoriesFilePath);
    const sheetName = workbook.SheetNames[0];
    const categoriesData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    // Assign a unique ID
    newCategory.id = categoriesData.length ? Math.max(...categoriesData.map(cat => cat.id)) + 1 : 1;

    categoriesData.push(newCategory);
    const newWorksheet = XLSX.utils.json_to_sheet(categoriesData);
    workbook.Sheets[sheetName] = newWorksheet;
    XLSX.writeFile(workbook, categoriesFilePath);
    res.json(categoriesData);
});



// Endpoint to add a new product
app.post('/products', (req, res) => {
    const newProduct = req.body;
    const workbook = XLSX.readFile(productsFilePath);
    const sheetName = workbook.SheetNames[0];
    const productsData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    // Assign a unique ID
    newProduct.id = productsData.length ? Math.max(...productsData.map(prod => prod.id)) + 1 : 1;

    productsData.push(newProduct);
    const newWorksheet = XLSX.utils.json_to_sheet(productsData);
    workbook.Sheets[sheetName] = newWorksheet;
    XLSX.writeFile(workbook, productsFilePath);
    res.json(productsData);
});

// Endpoint to get products by category ID
app.get('/products/category/:id', (req, res) => {
    const categoryId = parseInt(req.params.id);
    const workbook = XLSX.readFile(productsFilePath);
    const sheetName = workbook.SheetNames[0];
    const productsData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
    const filteredProducts = productsData.filter(product => product.category_id === categoryId);
    res.json(filteredProducts);
});

// Update Category
app.put('/categories/:id', (req, res) => {
    const { id } = req.params;
    const { name } = req.body;
    const workbook = XLSX.readFile(categoriesFilePath);
    const sheetName = workbook.SheetNames[0];
    const categories = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    const categoryIndex = categories.findIndex(cat => cat.id == id);
    if (categoryIndex >= 0) {
        categories[categoryIndex].name = name; // Update the category name
        const newSheet = XLSX.utils.json_to_sheet(categories);
        workbook.Sheets[sheetName] = newSheet;
        XLSX.writeFile(workbook, categoriesFilePath);
        res.json(categories[categoryIndex]);
    } else {
        res.status(404).json({ message: 'Category not found' });
    }
});

// Delete Category
app.delete('/categories/:id', (req, res) => {
    const { id } = req.params;
    const workbook = XLSX.readFile(categoriesFilePath);
    const sheetName = workbook.SheetNames[0];
    const categories = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    const updatedCategories = categories.filter(cat => cat.id != id);
    const newSheet = XLSX.utils.json_to_sheet(updatedCategories);
    workbook.Sheets[sheetName] = newSheet;
    XLSX.writeFile(workbook, categoriesFilePath);

    res.status(204).send(); // Send no content on successful deletion
});

// Function to read products from the Excel file
const readProductsFromExcel = () => {
    if (!fs.existsSync(productsFilePath)) {
        // If the file does not exist, return an empty array
        return [];
    }
    const workbook = XLSX.readFile(productsFilePath);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(worksheet);
    return data;
};

// Function to write products to the Excel file
const writeProductsToExcel = (products) => {
    const worksheet = XLSX.utils.json_to_sheet(products);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Products');
    XLSX.writeFile(workbook, productsFilePath);
};

// Endpoint to get all products
app.get('/products', (req, res) => {
    const products = readProductsFromExcel();
    res.json(products);
});

// Endpoint to update a product
app.put('/products/:id', (req, res) => {
    const { id } = req.params;
    const { name, price, category_id } = req.body;

    const products = readProductsFromExcel(); // Read current products
    const productIndex = products.findIndex(prod => prod.id === parseInt(id));

    if (productIndex > -1) {
        // Update the product details
        products[productIndex] = { id: parseInt(id), name, price, category_id };
        writeProductsToExcel(products); // Write updated products back to the Excel file
        return res.json(products[productIndex]); // Return the updated product
    } else {
        return res.status(404).json({ message: 'Product not found' });
    }
});

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
