const fs = require('fs');
const xlsx = require('xlsx');
const pptxgen = require('pptxgenjs');
const moment = require('moment');


// Ler os dados do Excel
const workbook = xlsx.readFile('dados_certificados.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = xlsx.utils.sheet_to_json(sheet);

// Modelo de texto do certificado
const modeloCertificado = "Certificamos para todos os fins que {NOME_DO_PARTICIPANTE}, portador(a) do CPF {CPF_PARTICIPANTE}, concluiu o curso {NOME_CURSO}, realizado no dia {DATA_CURSO}, com carga horária de {CARGA_HORARIA} horas e conteúdo programático relacionado no verso, promovido nas dependências da empresa {NOME_EMPRESA}, {ENDERECO_EMPRESA}.";

// Iterar sobre os dados e criar certificados em PowerPoint
data.forEach(row => {
    const nomeParticipante = row['Nome'];
    const cpfParticipante = row['CPF'];
    const nomeCurso = row['Nome do Curso'];
    const dataCurso = row['Data do Curso'] instanceof Date
        ? moment(row['Data do Curso']).format('DD/MM/YYYY')
        : row['Data do Curso'] ? xlsx.SSF.parse_date_code(row['Data do Curso']) : '';
    const cargaHoraria = row['Carga Horária'];
    const nomeEmpresa = row['Nome da Empresa'];
    const enderecoEmpresa = row['Endereço da Empresa'];

    // Criar uma apresentação em branco
    const presentation = new pptxgen();

    // Adicionar um slide com a imagem de fundo
    const slide = presentation.addSlide();
    slide.addImage({
        path: 'certificado.jpg', // Ajuste o caminho da imagem
        x: 0,
        y: 0,
        w: '100%', // Largura 100%
        h: '100%', // Altura 100%
        sizing: { type: 'contain', w: '100%', h: '100%' } // Ajusta a imagem para caber
    });

    // Substituir variáveis no modelo de certificado
    const textoCertificado = modeloCertificado
        .replace('{NOME_DO_PARTICIPANTE}', nomeParticipante)
        .replace('{CPF_PARTICIPANTE}', cpfParticipante)
        .replace('{NOME_CURSO}', nomeCurso)
        .replace('{DATA_CURSO}', dataCurso)
        .replace('{CARGA_HORARIA}', cargaHoraria)
        .replace('{NOME_EMPRESA}', nomeEmpresa)
        .replace('{ENDERECO_EMPRESA}', enderecoEmpresa);

    // Adicionar texto sobre a imagem
    slide.addText(textoCertificado, { x: '12%', y: '50%', fontFace: 'Arial', fontSize: 14, color: '363636', align: 'center', valign: 'middle' });

    // Salvar a apresentação
    const filePath = `Certificado_${nomeParticipante}.pptx`;
    presentation.writeFile(filePath, () => {
        console.log(`Arquivo salvo: ${filePath}`);
    });
});