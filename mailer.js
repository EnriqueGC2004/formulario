const express = require('express');
const bodyParser = require('body-parser');
const nodemailer = require('nodemailer');
const cors = require('cors');
const { Sequelize, DataTypes } = require('sequelize');

const app = express();
app.use(bodyParser.json());
app.use(cors());

// Configuración de la base de datos usando Sequelize
const sequelize = new Sequelize('uaem', 'root', '', {
  host: 'localhost',
  dialect: 'mysql',
});

// Definición del modelo Formulario
const Formulario = sequelize.define('formulario', {
  nombre: { type: DataTypes.STRING, allowNull: false },
  app: { type: DataTypes.STRING, allowNull: false },
  apm: { type: DataTypes.STRING, allowNull: false },
  numerodecuenta: { type: DataTypes.INTEGER, allowNull: false },
  dependencia: { type: DataTypes.STRING, allowNull: true },
  otro: { type: DataTypes.STRING, allowNull: true },
  grupo: { type: DataTypes.STRING, allowNull: false },
  correoelectronico: { type: DataTypes.STRING, allowNull: false },
  sexo: { type: DataTypes.STRING, allowNull: false },
  curso: { type: DataTypes.STRING, allowNull: false },
  cursoCategoria: { type: DataTypes.STRING, allowNull: false },
  instructor: { type: DataTypes.STRING, allowNull: false },
  correoelectronicoinstructor: { type: DataTypes.STRING, allowNull: false },
  rol: { type: DataTypes.STRING, allowNull: false },
  horario: { type: DataTypes.STRING, allowNull: false },
  fecha: { type: DataTypes.STRING, allowNull: false },
  duracion: { type: DataTypes.INTEGER, allowNull: false },
}, {
  tableName: 'formulario', // Nombre de la tabla en la base de datos
});

// Sincronizar el modelo con la base de datos
sequelize.sync()
  .then(() => console.log('Base de datos conectada y tabla sincronizada'))
  .catch(err => console.error('Error al conectar con MySQL:', err));

// Configuración de Nodemailer
const transporter = nodemailer.createTransport({
  host: 'smtp.gmail.com',
  port: 587,
  secure: false,
  auth: {
    user: 'al222211591@gmail.com',
    pass: 'vunk vozw pyjh zipm',
  },
});

async function generarYGuardarExcel() {
  try {
    const data = await Formulario.findAll();

    const ExcelJS = require('exceljs');
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Datos');

    worksheet.columns = [
      { header: 'Nombre', key: 'nombre', width: 20 },
      { header: 'Apellido Paterno', key: 'app', width: 20 },
      { header: 'Apellido Materno', key: 'apm', width: 20 },
      { header: 'Número de Cuenta', key: 'numerodecuenta', width: 20 },
      { header: 'Dependencia', key: 'dependencia', width: 20 },
      { header: 'Otro', key: 'otro', width: 20 },
      { header: 'Grupo', key: 'grupo', width: 20 },
      { header: 'Correo Electrónico', key: 'correoelectronico', width: 30 },
      { header: 'Sexo', key: 'sexo', width: 10 },
      { header: 'Curso Categoría', key: 'cursoCategoria', width: 20 },
      { header: 'Curso', key: 'curso', width: 20 },
      { header: 'Instructor', key: 'instructor', width: 20 },
      { header: 'Correo Electrónico del Instructor', key: 'correoelectronicoinstructor', width: 30 },
      { header: 'Rol', key: 'rol', width: 20 },
      { header: 'Horario', key: 'horario', width: 20 },
      { header: 'Fecha', key: 'fecha', width: 20 },
      { header: 'Duración (Horas)', key: 'duracion', width: 15 },
    ];

    data.forEach((record) => {
      worksheet.addRow(record.dataValues);
    });

    const filePath = 'C:\\Users\\kike2\\OneDrive\\Escritorio\\datos.xlsx';
    await workbook.xlsx.writeFile(filePath);

    console.log('Archivo Excel generado y guardado en:', filePath);
  } catch (error) {
    console.error('Error al generar y guardar el archivo Excel:', error);
  }
}

app.post('/api/formulario', async (req, res) => {
  try {
    const nuevoFormulario = await Formulario.create(req.body);

    await generarYGuardarExcel();

    const { nombre, app, apm, correoelectronico, curso, instructor, horario, fecha, duracion } = req.body;

    const nombreCompleto = `${nombre} ${app} ${apm}`;

    console.log(`Nombre completo: ${nombreCompleto}`);
    console.log(`Correo electrónico: ${correoelectronico}`);

    const output = `
      <html>
        <body style="margin: 0; padding: 0;">
          <table role="presentation" width="100%" style="border-collapse: collapse; background-image: url('https://your-domain.com/images/fondo.jpg'); background-size: cover; background-position: center; background-repeat: no-repeat;">
            <tr>
              <td align="center" style="padding: 20px;">
                <table role="presentation" width="600" style="background-color: rgba(0, 0, 0, 0.6); color: #fff; padding: 20px; border-radius: 8px;">
                  <tr>
                    <td>
                      <h1 style="font-size: 24px; font-weight: bold; margin: 0;">Hola ${nombreCompleto},</h1>
                      <p style="font-size: 16px; margin: 10px 0;">Tu inscripción al curso <i>${curso}</i> ha sido aceptada.</p>
                      <p style="font-size: 16px; margin: 10px 0;">Detalles del curso:</p>
                      <ul style="list-style-type: none; padding: 0; margin: 0;">
                        <li style="font-size: 16px; margin: 5px 0;"><strong>Instructor:</strong> ${instructor}</li>
                        <li style="font-size: 16px; margin: 5px 0;"><strong>Horario:</strong> ${horario}</li>
                        <li style="font-size: 16px; margin: 5px 0;"><strong>Fecha:</strong> ${fecha}</li>
                        <li style="font-size: 16px; margin: 5px 0;"><strong>Duración:</strong> ${duracion} horas</li>
                      </ul>
                      <p style="font-size: 16px; margin: 10px 0;">Gracias por inscribirte.</p>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>
        </body>
      </html>
    `;

    const mailOptions = {
      from: '"Cursos UAEM" <al222211591@gmail.com>',
      to: correoelectronico,
      subject: 'Confirmación de inscripción',
      html: output,
    };

    await transporter.sendMail(mailOptions);

    res.status(201).json(nuevoFormulario);
  } catch (error) {
    console.error('Error al registrar el usuario y enviar el correo:', error);
    res.status(400).json({ message: error.message });
  }
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
