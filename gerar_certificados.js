const fs = require('fs');
const xlsx = require('xlsx');
const pptxgen = require('pptxgenjs');
const moment = require('moment');



// Ler os dados do Excel
const workbook = xlsx.readFile('dados_certificados.xlsx', { cellText: false });
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = xlsx.utils.sheet_to_json(sheet);

const colunaCPF = sheet['B'];

// Inicializar um array para armazenar os CPFs
const cpfs = [];

// Iterar sobre as células da coluna "CPF" e extrair os valores
for (const celula in colunaCPF) {
    if (celula.startsWith('B') && colunaCPF.hasOwnProperty(celula)) {
        const valorCPF = colunaCPF[celula].v;
        // Adicionar o CPF ao array
        cpfs.push(valorCPF);
    }
}

function formatarCPF(cpf) {
    // Verificar se o CPF é uma string
    if (typeof cpf !== 'string') {
        // Se não for uma string, retornar o CPF original
        return cpf;
    }

    // Limpar CPF
    const cpfLimpo = cpf.replace(/\D/g, '');

    // Adicionar pontos e traço
    const cpfFormatado = cpfLimpo.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4');

    return cpfFormatado;
}


// Iterar sobre os CPFs e formatá-los
const cpfsFormatados = cpfs.map(formatarCPF);

// Agora, `cpfsFormatados` contém os CPFs formatados
console.log(cpfsFormatados);

// Modelo de texto do certificado
const modeloCertificado = "Certificamos para todos os fins que {NOME_DO_PARTICIPANTE}, portador(a) do CPF: {CPF_PARTICIPANTE}, concluiu o curso {NOME_CURSO}, realizado no dia {DATA_CURSO}, com carga horária de {CARGA_HORARIA} horas e conteúdo programático relacionado no verso, promovido nas dependências da empresa {NOME_EMPRESA}, {ENDERECO_EMPRESA}.";

// Iterar sobre os dados e criar certificados em PowerPoint
data.forEach(row => {
    const nomeParticipante = row['Nome'];
    const cpfParticipante = typeof row['CPF'] === 'string' ? formatarCPF(row['CPF'].trim()) : row['CPF'];
    const nomeCurso = row['Nome do Curso'];
    const cargaHoraria = row['Carga Horária'];
    const nomeEmpresa = row['Nome da Empresa'];
    const enderecoEmpresa = row['Endereço da Empresa'];
    const cpfParticipanteFormatado = formatarCPF(cpfParticipante);
    const dataCursoExcel = row['Data do Curso'];

    // Converter o número serial do Excel para data
    const dataCursoNumerica = parseFloat(dataCursoExcel);
    const dataCursoFormatada = !isNaN(dataCursoNumerica)
        ? moment("1900-01-01").add(dataCursoNumerica, 'days').format('DD/MM/YYYY')
        : 'data não informada';
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
        .replace('{CPF_PARTICIPANTE}', cpfParticipanteFormatado)
        .replace('{NOME_CURSO}', nomeCurso)
        .replace('{DATA_CURSO}', dataCursoFormatada)
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
