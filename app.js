const express = require('express');
const path = require('path');
const mongoose = require('mongoose');
const bodyParser = require('body-parser');
const exceljs = require('exceljs');
const ExcelJS = require('exceljs');


const app = express();

mongoose.connect('mongodb+srv://velammal:1234@nodeexpressprojects.sruvfla.mongodb.net/BOOK-WEBSITE?retryWrites=true&w=majority', {
  useNewUrlParser: true,
  useUnifiedTopology: true,
});

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));


const bookSchema = new mongoose.Schema({
  name: String,
  title: String,
  authorPosition: String,
  department:String,  
  publisherName: String,
  publicationDate: Date,
  image: String,
  remarks: String
});

const Book = mongoose.model('Book', bookSchema);

const chapterSchema = new mongoose.Schema({
  name: String,
  title: String,
  titleofchapter:String,
  authorPosition: String,
  department:String,
  publicationDate: Date,
  publisherName: String,
  frompg:String,
  topg : String,
  image:String,
  remarks: String,
  bookType: String,
});

const Chapter = mongoose.model('Chapter', bookSchema);



const staffSchema = new mongoose.Schema({
  name: String,
  email: String,
  password: String,
});


const Staff = mongoose.model('Staff', staffSchema);

const adminSchema = new mongoose.Schema({
  admin_id: String, 
  email: String,
  password: String,
});

const Admin = mongoose.model('Admin', adminSchema, 'admins');



app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'home.html'));
});

app.get('/staff/signup', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'staffsignup.html'));
});

app.get('/staff/signin', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'staffsignin.html'));
});
app.get('/bookorchapter', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'BookorChapter.html'));
});
app.get('/chapterupload', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'chapterupload.html'));
});
app.get('/bookupload', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'bookupload.html'));
});

app.get('/admin/signin', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'adminsignin.html'));
});
app.get('/admin', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'admin.html'));
});



app.get('/admin/download/books', async (req, res) => {
  try {
   
    const books = await Book.find();

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Books');

    
    worksheet.columns = [
      { header: 'Name', key: 'name'},
      { header: 'Title', key: 'title'},
      { header: 'Author Position', key: 'authorPosition'},
      { header: 'Publication Date', key: 'publicationDate'},
      { header: 'Publisher Name', key: 'publisherName'},
      { header: 'PDF Link', key: 'url'},
      { header: 'Remarks', key: 'remarks'},
      { header: 'Book Type', key: 'bookType'},
     
    ];

   
    books.forEach(book => {
      worksheet.addRow({
        name: book.name,
        title: book.title,
        authorPosition: book.authorPosition,
        publicationDate: book.publicationDate,
        publisherName: book.publisherName,
        url: book.url,
        remarks: book.remarks,  
        bookType: book.bookType, 
        
      });
    });

  
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=books.xlsx');

   
    workbook.xlsx.write(res);

  } catch (error) {
    console.error(error);
    res.status(500).send('Internal Server Error');
  }
});

app.get('/admin/preview/books', async (req, res) => {
  try {
    const books = await Book.find(); 

    res.render('preview-books', { books }); 
  } catch (error) {
    console.error(error);
    res.status(500).send('Internal Server Error');
  }
});

app.post('/staff/signup', async (req, res) => {
  try {
    const { name, email, password } = req.body;

    const emailRegex = /^[a-zA-Z0-9._%+-]+@velammalit\.edu\.in$/;
    if (!emailRegex.test(email)) {
      return res.status(400).json({ error: 'Invalid email format' });
    }

    const existingStaff = await Staff.findOne({ email });
    if (existingStaff) {
      return res.status(409).json({ error: 'Email already registered' });
    }

    const newStaff = new Staff({ name, email, password });
    await newStaff.save();

    res.redirect('/'); 
  } catch (error) {
    console.error(error);
    res.status(500).send('Internal Server Error');
  }
});

app.post('/staff/signin', async (req, res) => {
  try {
    const { email, password } = req.body;

    const staff = await Staff.findOne({ email, password });
    if (!staff) {
      return res.status(401).send('Invalid credentials');
    }  
    res.redirect('/bookorchapter');

  } catch (error) {
    console.error(error);
    res.status(500).send('Internal Server Error');
  }
});


app.post('/bookupload', async (req, res) => {
  try {
    const { name, title, authorPosition, department, publicationDate, publisherName, image, remarks} = req.body;

  
    const newBook = new Book({
      name,
      title,
      authorPosition,
      department,      
      publisherName,
      publicationDate,
      image,
      remarks
    });

   
    await newBook.save();

    res.redirect('/bookorchapter'); 
  } catch (error) {
    console.error(error);
    res.status(500).send('Internal Server Error');
  }
});

app.post('/chapterupload', async (req, res) => {
  try {
    const { name, title, titleofchapter, authorPosition, department,publicationname, frompg, topg,publicationDate,image,remarks } = req.body;
 
    const newChapter = new Chapter({
      name,
      title,
      titleofchapter,
      authorPosition,
      department,
      publicationname,
      frompg,
      topg,
      publicationDate,
      image,
      remarks
    });

   
    await newChapter.save();

    res.redirect('/bookorchapter'); 
  } catch (error) {
    console.error(error);
    res.status(500).send('Internal Server Error');
  }
});

app.post('/admin/signin', async (req, res) => {
  try {
    const { admin_id, email, password } = req.body;

   
    const admin = await Admin.findOne({ admin_id });
    if (!admin) {
      return res.status(401).send('Admin ID not found');
    }

    if (admin.email !== email || admin.password !== password) {
      return res.status(401).send('Invalid credentials');
    }
    res.redirect('/admin'); 

  } catch (error) {
    console.error(error);
    res.status(500).send('Internal Server Error');
  }
});


const db = mongoose.connection;
db.on('error', (error) => {
  console.error('MongoDB connection error:', error);
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
