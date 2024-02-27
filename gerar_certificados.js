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

    // Formatação CPF
    const cpfLimpo = cpf.replace(/\D/g, '');

    // Adicionar pontos e traço
    const cpfFormatado = cpfLimpo.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4');

    return cpfFormatado;
}


// Iterar sobre os CPFs e formatá-los
const cpfsFormatados = cpfs.map(formatarCPF);
console.log(cpfsFormatados);

// Modelo de texto do certificado
const modeloCertificado = "Certifico que {NOME_DO_PARTICIPANTE}, portador(a) do CPF: {CPF_PARTICIPANTE}, patrocinado pela {NOME_EMPRESA}, participou e foi aprovado no treinamento de {NOME_CURSO},  com carga horária de {CARGA_HORARIA}, realizado no dia {DATA_CURSO}, ministrado pelo Engenheiro de Segurança do Trabalho, {ENGENHEIRO}, CREA/ SP – 5.061.417.650 .";
const dataEstado = "São Paulo, {DATA_CURSO} ";

// Iterar sobre os dados e criar certificados em PowerPoint
data.forEach(row => {
     const presentation = new pptxgen();
    const nomeParticipante = row['Nome'];
    const cpfParticipante = typeof row['CPF'] === 'string' ? formatarCPF(row['CPF'].trim()) : row['CPF'];
    const nomeCurso = row['Nome do Curso'];
    const cargaHoraria = row['Carga Horária'];
    const nomeEmpresa = row['Nome da Empresa'];
    const Engenheiro = row['Engenheiro'];
    const cpfParticipanteFormatado = formatarCPF(cpfParticipante);
    const dataCursoExcel = row['Data do Curso'];

    // Converter o número serial do Excel para data
    const dataCursoNumerica = parseFloat(dataCursoExcel);
    const dataCursoFormatada = !isNaN(dataCursoNumerica)
        ? moment("1900-01-01").add(dataCursoNumerica, 'days').format('DD/MM/YYYY')
        : 'data não informada';
    // Criar uma apresentação em branco
   

    // Adicionar um slide com a imagem de fundo
    const slide = presentation.addSlide();
    slide.addImage({
        path: 'CertificadosNR35.png', 
        x: 0,
        y: 0,
        w: '100%',
        h: '100%', 
        sizing: { type: 'contain', w: '100%', h: '100%' } 
    });

    // Substituir as varaiaveis criadas
    const textoCertificado = modeloCertificado
        .replace('{NOME_DO_PARTICIPANTE}', nomeParticipante)
        .replace('{CPF_PARTICIPANTE}', cpfParticipanteFormatado)
        .replace('{NOME_CURSO}', nomeCurso)
        .replace('{DATA_CURSO}', dataCursoFormatada)
        .replace('{CARGA_HORARIA}', cargaHoraria)
        .replace('{NOME_EMPRESA}', nomeEmpresa)
        .replace('{ENGENHEIRO}', Engenheiro);

    //Ajustar o texto e a fonte correta
    slide.addText(textoCertificado, { x: '12%', y: '43%', fontFace: 'Arial', fontSize: 13, color: '363636', align: 'left', valign: 'middle', charSpacing: 1 });

    const textoSimples = dataEstado
    .replace('{DATA_CURSO}', dataCursoFormatada);
    slide.addText(textoSimples, { x: '37%', y: '57%', fontFace: 'Arial', fontSize: 12.5, color: '363636', align: 'center', valign: 'middle', charSpacing: 1 });

    const slideConteudoProgramatico = presentation.addSlide();
    slideConteudoProgramatico.addImage({
        path: 'NR35_ContProg.png',  
        x: 0, 
        y: 0, 
        w: '100%', 
        h: '100%', 
        sizing: { type: 'contain', w: '100%', h: '100%' }})
    // Salvar o certificado
    const filePath = `Certificado_${nomeParticipante}.pptx`;
    presentation.writeFile(filePath, () => {
        console.log(`Arquivo salvo: ${filePath}`);
    });
});
