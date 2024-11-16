const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const port = 3000;

// Criar o diretório 'uploads' caso não exista
const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) {
    fs.mkdirSync(uploadDir, { recursive: true });
}

// Configuração do multer para salvar o arquivo temporariamente
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, uploadDir);  // Salvar na pasta 'uploads'
    },
    filename: (req, file, cb) => {
        cb(null, Date.now() + path.extname(file.originalname)); // Garantir nome único para o arquivo
    }
});
const upload = multer({ storage: storage });

// Rota para exibir o formulário HTML
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Rota para processar o upload e conversão
app.post('/upload', upload.single('xlsFile'), (req, res) => {
    const xlsFilePath = req.file.path;
    const xlsxFilePath = path.join(__dirname, 'converted', Date.now() + '.xlsx');
    
    try {
        // Ler o arquivo XLS
        const workbook = XLSX.readFile(xlsFilePath);
        
        // Combinar todas as planilhas em uma única planilha
        let combinedData = [];
        
        // Iterar por todas as planilhas
        workbook.SheetNames.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });  // Converte a planilha para um array de arrays (linhas e colunas)
            
            // Adicionar os dados dessa planilha ao conjunto combinado
            combinedData = combinedData.concat(data);
        });

        // Criar uma nova planilha com os dados combinados
        const newWorksheet = XLSX.utils.aoa_to_sheet(combinedData);  // Converte o array de arrays para uma planilha
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'CombinedSheet');  // Adiciona a planilha combinada ao novo workbook

        // Criar a pasta 'converted' se não existir
        const convertedDir = path.join(__dirname, 'converted');
        if (!fs.existsSync(convertedDir)) {
            fs.mkdirSync(convertedDir, { recursive: true });
        }

        // Escrever o arquivo XLSX
        XLSX.writeFile(newWorkbook, xlsxFilePath);

        // Apagar o arquivo XLS temporário
        fs.unlinkSync(xlsFilePath);

        // Retornar o link para o arquivo XLSX convertido
        res.send(`<h1>Arquivo convertido com sucesso!</h1><p><a href="/download/${path.basename(xlsxFilePath)}">Baixar arquivo convertido</a></p>`);
    } catch (error) {
        console.error(error);
        res.status(500).send('Erro ao processar o arquivo.');
    }
});

// Rota para baixar o arquivo XLSX convertido
app.get('/download/:filename', (req, res) => {
    const filePath = path.join(__dirname, 'converted', req.params.filename);
    res.download(filePath);
});

app.listen(port, () => {
    console.log(`Servidor rodando em http://localhost:${port}`);
});
