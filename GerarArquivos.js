async function loadImageAsBase64(url)
{
    const resp = await fetch(url);
    const blob = await resp.blob();
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
}

async function exportJsonToExcel(data) {
  // Create a new workbook
  const workbook = XLSX.utils.book_new();

    if (!Array.isArray(data)) {
        data = [data]; // transforma em array com um único objeto, se necessário
    }
  // Convert JSON data to a worksheet
  const worksheet = XLSX.utils.json_to_sheet(data);

  // Append the worksheet to the workbook
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

  // Configurar data e escrever tabela xlsx
  const dataAtual = new Date();
  const dia = String(dataAtual.getDate()).padStart(2, '0'); // Adiciona zero à esquerda se necessário
  const mes = String(dataAtual.getMonth() + 1).padStart(2, '0'); // +1 porque meses são 0-11
  const ano = dataAtual.getFullYear();
  const hora = dataAtual.getHours();
  const minto = dataAtual.getMinutes();
  const sgdo = dataAtual.getSeconds();
  XLSX.writeFile(workbook, `Relatorio Completo - ${dia}${mes}${ano}_${hora}${minto}${sgdo}.xlsx`);
}


async function readExcelFile(file) {
  const XLSX = window.XLSX;

  if (!XLSX) {
    console.error("A biblioteca XLSX não está carregada. Certifique-se de incluir o script do SheetJS.");
    return null;
  }

  try {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Extrair cabeçalhos e dados
    const headers = jsonData[0] || [];
    const rows = jsonData.slice(1).map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index] !== undefined ? row[index] : '';
      });
      return obj;
    });

    return rows;
  } catch (error) {
    console.error("Erro ao ler o arquivo Excel:", error);
    return null;
  }
}

async function mergeExcelData(files) {
  const XLSX = window.XLSX;

  if (!XLSX) {
    console.error("A biblioteca XLSX não está carregada. Certifique-se de incluir o script do SheetJS.");
    return;
  }

  if (files.length < 2) {
    alert("Por favor, selecione pelo menos dois arquivos Excel.");
    return;
  }

  const combinedData = [];
  let headers = [];

  for (const file of files) {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    if (headers.length === 0) {
      headers = jsonData[0] || [];
    } else if (jsonData[0] && JSON.stringify(jsonData[0]) !== JSON.stringify(headers)) {
      alert("Os cabeçalhos dos arquivos não são iguais. A concatenação pode não funcionar corretamente.");
      return;
    }

    const rows = jsonData.slice(1).map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index] !== undefined ? row[index] : '';
      });
      return obj;
    });

    combinedData.push(...rows);
  }

  const worksheetData = [headers, ...combinedData.map(item => headers.map(header => item[header] || ''))];
  const newWorkbook = XLSX.utils.book_new();
  const newWorksheet = XLSX.utils.aoa_to_sheet(worksheetData);
  XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'CombinedSheet');

  const dataAtual = new Date();
  const dia = String(dataAtual.getDate()).padStart(2, '0');
  const mes = String(dataAtual.getMonth() + 1).padStart(2, '0');
  const ano = dataAtual.getFullYear();
  const hora = dataAtual.getHours();
  const minto = dataAtual.getMinutes();
  const sgdo = dataAtual.getSeconds();
  XLSX.writeFile(newWorkbook, `Relatorio Combinado - ${dia}${mes}${ano}_${hora}${minto}${sgdo}.xlsx`);
}


async function gerarGuiaDeSepultamentoPDF(data) {
    const {jsPDF} = window.jspdf;
    const doc = new jsPDF();
    //falta adicionar o campo de Data
    const logoSDUurl = "https://i.imgur.com/qheEa6L.png"; // upload do brasao para o Imgur
    const logoSDUwidth = 65;
    const logoSDUheight = logoSDUwidth * 451 / 2173;
    const logoSDUx = 10;
    const logoSDUy = 10;
    const imageBase64SDU = await loadImageAsBase64(logoSDUurl);
    doc.addImage(imageBase64SDU, 'PNG', logoSDUx, logoSDUy, logoSDUwidth, logoSDUheight);


    const logoPMTurl = "https://i.imgur.com/jkAUcJQ.png"; // upload do brasao para o Imgur
    const logoPMTwidth = 65;
    const logoPMTheight = logoPMTwidth * 653 / 2567;
    const logoPMTx = logoSDUx + logoSDUwidth + 5;
    const logoPMTy = 10;
    const imageBase64PMT = await loadImageAsBase64(logoPMTurl);
    doc.addImage(imageBase64PMT, 'PNG', logoPMTx, logoPMTy, logoPMTwidth, logoPMTheight);


    let y = 50;
    doc.setFontSize(14);
    doc.text("GUIA DE SEPULTAMENTO", 70, y);
    y += 15;

    doc.setFontSize(10);

    // Linha 1: DECLARAÇÃO / REGISTRO
    doc.text("DECLARAÇÃO DE ÓBITO Nº", 10, y);
    doc.text("REGISTRO DE ÓBITO Nº", 110, y);
    y += 4;
    doc.rect(10, y, 80, 8); //(x,y, largura, altura)
    doc.rect(110, y, 80, 8);
    y += 15;

    // Linha 2: LOCAL / FETAL / INDIGENTE
    doc.text("LOCAL DE ÓBITO", 10, y);
    doc.text("ÓBITO FETAL", 140, y);
    doc.text("INDIGENTE", 170, y);
    y += 4;
    doc.rect(10, y, 120, 8);
    doc.rect(140, y, 20, 8);
    doc.rect(170, y, 20, 8);
    y += 15;

    // NOME
    doc.text("NOME", 10, y);
    y += 4;
    doc.rect(10, y, 180, 8);
    y += 15;

    // PAI
    doc.text("PAI", 10, y);
    y += 4;
    doc.rect(10, y, 180, 8);
    y += 15;

    // MÃE
    doc.text("MÃE", 10, y);
    y += 4;
    doc.rect(10, y, 180, 8);
    y += 15;

    // FUNERÁRIA
    doc.text("FUNERÁRIA", 10, y);
    y += 4;
    doc.rect(10, y, 180, 8);
    y += 15;

    // NOTA FISCAL
    doc.text("NOTA FISCAL", 10, y);
    y += 4;
    doc.rect(10, y, 180, 8);
    y += 15;

    // CEMITÉRIO
    doc.text("CEMITÉRIO", 10, y);
    y += 4;
    doc.rect(10, y, 180, 8);
    y += 15;

    // SEÇÃO, QUADRA, FILA, COVA, PLACA
    doc.text("SEÇÃO", 10, y);
    doc.text("QUADRA", 50, y);
    doc.text("FILA", 90, y);
    doc.text("COVA", 130, y);
    doc.text("PLACA", 170, y);
    y += 4;
    doc.rect(10, y, 30, 8);
    doc.rect(50, y, 30, 8);
    doc.rect(90, y, 30, 8);
    doc.rect(130, y, 30, 8);
    doc.rect(170, y, 30, 8);

    y += 15;
    doc.text("HORÁRIO DO SEPULTAMENTO", 10, y);
    y += 8;
    doc.text("PREVISTO PARA:", 10, y);
    y += 4;
    doc.rect(10, y, 150, 8);

    y += 15;
    doc.text("EXPEDIÇÃO", 10, y);
    y += 4;
    doc.rect(10, y, 90, 8);
    doc.text("VISTO", 120, y);//O que entra aqui?


    if (y > 700) {
        doc.addPage();
        y = 20;
    }
    doc.autoTable({
        startY: y + 10,
        head: [['Taxas', 'Valor']],
        body: [
            ['Abertura de Sepultura', 'R$0,00'],
            ['Reabertura de Sepultura', 'R$0,00'],
            ['Inumamação Adulto', 'R$0,00'],
            ['Inumação Infantil', 'R$0,00'],
        ],
        tableWidth: 'wrap',
        headStyles: {
            fillColor: [240, 240, 240], //cinza claro
            textColor: 0, // Texto preto
            halign: 'left',
            valign: 'middle',
            fontSize: 14,
        },
        styles: {
            halign: 'left',
            valign: 'middle',
            fontSize: 14,
        },
        columnStyles: {
            0: {fillColor: [240, 240, 240]},  // Cinza claro
            1: {fillColor: false}             // Sem cor de fundo
        }
    });
    // Exporta o PDF
    doc.save("guia_sepultamento.pdf");
}

async function gerarRelatorioFmsPDF(data) {
    const { jsPDF } = window.jspdf;
    const dataAtual = new Date();
    const dia = String(dataAtual.getDate()).padStart(2, '0');
    const mes = String(dataAtual.getMonth() + 1).padStart(2, '0');
    const ano = dataAtual.getFullYear();
    const hora = dataAtual.getHours();
    const minto = dataAtual.getMinutes();
    const sgdo = dataAtual.getSeconds();

    const doc = new jsPDF({
        orientation: "landscape",
        unit: "mm",
        format: "a4",
    });

    const logoUrl = "https://i.imgur.com/vaDmtrQ.png";
    const logoWidth = 25;
    const logoHeight = logoWidth * 1318 / 1200;
    const logoX = 10;
    const logoY = 10;
    const imageBase64 = await loadImageAsBase64(logoUrl);
    doc.addImage(imageBase64, 'PNG', logoX, logoY, logoWidth, logoHeight);

    const distancia_logo_texto = 5;
    const texto_cabecalhoX = logoX + logoWidth + distancia_logo_texto;
    const texto_cabecalhoY = logoY + 5;

    doc.setFontSize(11);
    doc.setFont(undefined, 'bold');
    doc.text("ESTADO DO PIAUÍ", texto_cabecalhoX, texto_cabecalhoY);
    doc.text("Prefeitura Municipal de Teresina", texto_cabecalhoX, texto_cabecalhoY + 6);
    doc.text("Superintendência de Desenvolvimento Urbano - Centro (SDU-Centro)", texto_cabecalhoX, texto_cabecalhoY + 12);
    doc.text("Plantão Funerário", texto_cabecalhoX, texto_cabecalhoY + 18);

    const titulo = `RELATÓRIO - ${dia}/${mes}/${ano}`;
    doc.setFontSize(16);
    doc.setFont(undefined, 'bold');
    const pageWidth = doc.internal.pageSize.getWidth();
    const titleWidth = doc.getTextWidth(titulo);
    doc.text(titulo, (pageWidth - titleWidth) / 2, 60);
    doc.setFont(undefined, 'normal');

    const calcularIdade = (dataNascimento) => {
        if (!dataNascimento) return 'N/A';
        const [diaNasc, mesNasc, anoNasc] = dataNascimento.split('/').map(Number);
        const nascimento = new Date(anoNasc, mesNasc - 1, diaNasc);
        const diff = new Date() - nascimento;
        const idade = Math.floor(diff / (1000 * 60 * 60 * 24 * 365.25));
        return idade;
    };

    // Iterar sobre o array de dados
    doc.autoTable({
        startY: 70,
        head: [["ITEM", "NOME", "DATA DO \nNASCIMENTO", "IDADE", "BAIRRO", "MUNICÍPIO", "DATA DO \nFALECIMENTO", "LOCAL DO FALECIMENTO", "CAUSA DA MORTE", "CEMITÉRIO"]],
        body: data.map((item, index) => [
            (index + 1).toString(),
            item.nome || "Não informado",
            item.dataNascimento || "Não informado",
            item.dataNascimento ? calcularIdade(item.dataNascimento) : 'N/A',
            item.bairro || "Não informado",
            item.municipio || "Teresina",
            item.dataObito || "Não informado",
            item.localObito || "Não informado",
            item.causaMorte || "Não informado",
            item.cemiterio || "Não informado"
        ]),
        styles: {
            fontSize: 8,
            valign: 'middle',
            halign: 'center',
            cellPadding: 1,
            overflow: 'linebreak'
        },
        columnStyles: {
            0: { cellWidth: 15 },
            1: { cellWidth: 30 },
            2: { cellWidth: 30 },
            3: { cellWidth: 15 },
            4: { cellWidth: 25 },
            5: { cellWidth: 25 },
            6: { cellWidth: 30 },
            7: { cellWidth: 30 },
            8: { cellWidth: 25 },
            9: { cellWidth: 30 }
        },
        headStyles: {
            fillColor: [240, 240, 240],
            textColor: [0, 0, 0],
            fontSize: 10,
            fontStyle: 'bold',
            halign: 'center',
            valign: 'middle'
        },
        theme: 'grid',
        margin: { left: 10, right: 10 }
    });

    doc.save(`Relatorio__${dia}${mes}${ano}_${hora}${minto}${sgdo}.pdf`);
}

async function gerarRelatorioSduPDF(data) {
    const { jsPDF } = window.jspdf;
    const dataAtual = new Date();
    const dia = String(dataAtual.getDate()).padStart(2, '0');
    const mes = String(dataAtual.getMonth() + 1).padStart(2, '0');
    const ano = dataAtual.getFullYear();
    const hora = dataAtual.getHours();
    const minto = dataAtual.getMinutes();
    const sgdo = dataAtual.getSeconds();

    const doc = new jsPDF({
        orientation: "landscape",
        unit: "mm",
        format: "a4",
    });

    const logoUrl = "https://i.imgur.com/vaDmtrQ.png";
    const logoWidth = 25;
    const logoHeight = logoWidth * 1318 / 1200;
    const logoX = 10;
    const logoY = 10;
    const imageBase64 = await loadImageAsBase64(logoUrl);
    doc.addImage(imageBase64, 'PNG', logoX, logoY, logoWidth, logoHeight);

    const distancia_logo_texto = 5;
    const texto_cabecalhoX = logoX + logoWidth + distancia_logo_texto;
    const texto_cabecalhoY = logoY + 5;

    doc.setFontSize(11);
    doc.setFont(undefined, 'bold');
    doc.text("ESTADO DO PIAUÍ", texto_cabecalhoX, texto_cabecalhoY);
    doc.text("Prefeitura Municipal de Teresina", texto_cabecalhoX, texto_cabecalhoY + 6);
    doc.text("Superintendência de Desenvolvimento Urbano - Centro (SDU-Centro)", texto_cabecalhoX, texto_cabecalhoY + 12);
    doc.text("Plantão Funerário", texto_cabecalhoX, texto_cabecalhoY + 18);

    const titulo = `RELATÓRIO DESCRITIVO SDU-CENTRO - ${dia}/${mes}/${ano}`;
    doc.setFontSize(16);
    doc.setFont(undefined, 'bold');
    const pageWidth = doc.internal.pageSize.getWidth();
    const titleWidth = doc.getTextWidth(titulo);
    doc.text(titulo, (pageWidth - titleWidth) / 2, 60);
    doc.setFont(undefined, 'normal');

    doc.autoTable({
        startY: 70,
        head: [["DECLARAÇÃO DE ÓBITO", "FUNERÁRIA", "CEMITÉRIO", "BLOCO", "GUIA", "TIPO DE TAXA", "VALOR TAXA", "DATA DE ATENDIMENTO", "PLANTONISTA"]],
        body: data.map((item, index) => [
            item.declaracaoObito || "Não informado",
            item.empresa || "Não informado",
            item.cemiterio || "Não informado",
            item.bairro || "Não informado", // Ajuste conforme necessário
            item.guia || "Não informado",
            item.tipoTaxa || "Não informado",
            item.valorTaxa || "Não informado",
            item.dataAtendimento || "Não informado",
            item.plantonista || "Não informado"
        ]),
        styles: {
            fontSize: 8,
            valign: 'middle',
            halign: 'center',
            cellPadding: 1,
            overflow: 'linebreak'
        },
        columnStyles: {
            0: { cellWidth: 30 },
            1: { cellWidth: 25 },
            2: { cellWidth: 25 },
            3: { cellWidth: 20 },
            4: { cellWidth: 20 },
            5: { cellWidth: 25 },
            6: { cellWidth: 20 },
            7: { cellWidth: 30 },
            8: { cellWidth: 28 }
        },
        headStyles: {
            fillColor: [240, 240, 240],
            textColor: [0, 0, 0],
            fontSize: 10,
            fontStyle: 'bold',
            halign: 'center',
            valign: 'middle'
        },
        theme: 'grid',
        margin: { left: 30, right: 10 }
    });

    // Adicionar nova página para análise
    doc.addPage();
    doc.setFontSize(16);
    doc.setFont(undefined, 'bold');

    doc.addImage(imageBase64, 'PNG', logoX, logoY, logoWidth, logoHeight);
    doc.text("ESTADO DO PIAUÍ", texto_cabecalhoX, texto_cabecalhoY);
    doc.text("Prefeitura Municipal de Teresina", texto_cabecalhoX, texto_cabecalhoY + 6);
    doc.text("Superintendência de Desenvolvimento Urbano - Centro (SDU-Centro)", texto_cabecalhoX, texto_cabecalhoY + 12);
    doc.text("Plantão Funerário", texto_cabecalhoX, texto_cabecalhoY + 18);

    const estatisticoTitulo = "RELATÓRIO ESTATÍSTICO SDU-CENTRO";
    const estatisticoTitleWidth = doc.getTextWidth(estatisticoTitulo);
    doc.text(estatisticoTitulo, (pageWidth - estatisticoTitleWidth) / 2, 20);

    // Calcular estatísticas
    const variacoesSaoJose = ["são josé", "sao jose", "são jose", "sao josé"];
    const variacoesSaoJudasTadeu = ["são judas tadeu", "sao judas tadeu"];
    const variacoesSantoAntonio = ["santo antônio", "santo antonio"]
    const variacoesRenascenca = ["renascença", "renascenca"]
    const variacoesSaoSebastiao = ["são sebastião", "são sebastiao", "sao sebastião", "sao sebastiao"]
    const variacoesSaoJoaoBatista = ["são joão batista", "são joao batista", "sao joão batista", "sao joao batista"]
    const variacoesSantaMonica = ["santa mônica", "santa monica"]

    const saoJoseCount = data.filter(item =>
        item.cemiterio &&
        variacoesSaoJose.some(variacao =>
            item.cemiterio.toLowerCase() === variacao
        )
    ).length;

    const saoJudasTadeuCount = data.filter(item =>
        item.cemiterio &&
        variacoesSaoJudasTadeu.some(variacao =>
            item.cemiterio.toLowerCase() === variacao
        )
    ).length;

    const santoAntonioCount = data.filter(item =>
        item.cemiterio &&
        variacoesSantoAntonio.some(variacao =>
            item.cemiterio.toLowerCase() === variacao
        )
    ).length;

    const renascencaCount = data.filter(item =>
        item.cemiterio &&
        variacoesRenascenca.some(variacao =>
            item.cemiterio.toLowerCase() === variacao
        )
    ).length;

    const saoSebastiaoCount = data.filter(item =>
        item.cemiterio &&
        variacoesSaoSebastiao.some(variacao =>
            item.cemiterio.toLowerCase() === variacao
        )
    ).length;

    const saoJoaoBatistaCount = data.filter(item =>
        item.cemiterio &&
        variacoesSaoJoaoBatista.some(variacao =>
            item.cemiterio.toLowerCase() === variacao
        )
    ).length;

    const santaMonicaCount = data.filter(item =>
        item.cemiterio &&
        variacoesSantaMonica.some(variacao =>
            item.cemiterio.toLowerCase() === variacao
        )
    ).length;

    const totalTaxasSaoJose = 0;

    // Tabela de análise
    doc.autoTable({
        startY: 70,
        head: [["Cemitério", "Jazigos", "Quantidade de isentos", "Valor total de taxas"]],
        body: [
            ["São José", saoJoseCount],
            ["São Judas Tadeu", saoJudasTadeuCount],
            ["Dom Bosco", domBoscoCount],
            ["Santo Antônio", santoAntonioCount],
            ["Renascença", renascencaCount],
            ["Poty Velho", potyVelhoCount],
            ["São Sebastião", saoSebastiaoCount],
            ["São João Batista", saoJoaoBatistaCount],
            ["Morros", morrosCount],
            ["Santa Cruz", santaCruzCount],
            ["Santa Mônica", santaMonicaCount]
        ],
        styles: {
            fontSize: 10,
            valign: 'middle',
            halign: 'left',
            cellPadding: 2,
            overflow: 'linebreak'
        },
        columnStyles: {
            0: { cellWidth: 80 },
            1: { cellWidth: 50, halign: 'right' }
        },
        headStyles: {
            fillColor: [240, 240, 240],
            textColor: [0, 0, 0],
            fontSize: 12,
            fontStyle: 'bold',
            halign: 'center',
            valign: 'middle'
        },
        theme: 'grid',
        margin: { left: 20, right: 10 }
    });



    doc.autoTable({
        startY: 140,
        head: [["Cemitério (Zonas)", "Jazigos", "Quantidade de isentos", "Valor total de taxas"]],
        body: [
            ["Norte", saoJoseCount],
            ["Leste", saoJudasTadeuCount],
            ["Sul", domBoscoCount],
            ["Sudeste", santoAntonioCount]
        ],
        styles: {
            fontSize: 10,
            valign: 'middle',
            halign: 'left',
            cellPadding: 2,
            overflow: 'linebreak'
        },
        columnStyles: {
            0: { cellWidth: 80 },
            1: { cellWidth: 50, halign: 'right' }
        },
        headStyles: {
            fillColor: [240, 240, 240],
            textColor: [0, 0, 0],
            fontSize: 12,
            fontStyle: 'bold',
            halign: 'center',
            valign: 'middle'
        },
        theme: 'grid',
        margin: { left: 20, right: 10 }
    });

    doc.save(`Relatorio_SDU_${dia}${mes}${ano}_${hora}${minto}${sgdo}.pdf`);
}


async function gerarRelatorioCompletoPDF(data) {
    const {jsPDF} = window.jspdf;
    const dataAtual = new Date();
    const dia = String(dataAtual.getDate()).padStart(2, '0'); // Adiciona zero à esquerda se necessário
    const mes = String(dataAtual.getMonth() + 1).padStart(2, '0'); // +1 porque meses são 0-11
    const ano = dataAtual.getFullYear();

    const doc = new jsPDF({
        orientation: "landscape",
        unit: "mm",
        format: "a4",
    });

    const logoUrl = "https://i.imgur.com/vaDmtrQ.png"; // upload do brasao para o Imgur
    const logoWidth = 25;
    const logoHeight = logoWidth * 1318 / 1200;
    const logoX = 10;
    const logoY = 10;
    const imageBase64 = await loadImageAsBase64(logoUrl);
    doc.addImage(imageBase64, 'PNG', logoX, logoY, logoWidth, logoHeight);

    const distancia_logo_texto = 5;
    const texto_cabecalhoX = logoX + logoWidth + distancia_logo_texto;
    const texto_cabecalhoY = logoY + 5;

    doc.setFontSize(11);
    doc.setFont(undefined, 'bold');
    doc.text("ESTADO DO PIAUÍ", texto_cabecalhoX, texto_cabecalhoY);
    doc.text("Prefeitura Municipal de Teresina", texto_cabecalhoX, texto_cabecalhoY + 6);
    doc.text("Superintendência de Desenvolvimento Urbano - Centro (SDU-Centro)", texto_cabecalhoX, texto_cabecalhoY + 12);
    doc.text("Plantão Funerário", texto_cabecalhoX, texto_cabecalhoY + 18);


    const titulo = `RELATÓRIO - ${dia}/${mes}/${ano}`;
    // Set font size and style
    doc.setFontSize(16);
    doc.setFont(undefined, 'bold');

    const pageWidth = doc.internal.pageSize.getWidth();
    const titleWidth = doc.getTextWidth(titulo);
    doc.text(titulo, (pageWidth - titleWidth) / 2, 60);
    doc.setFont(undefined, 'normal');

    const calcularIdade = (dataNascimento) => {
        if (!dataNascimento) return 'N/A';
        const [diaNasc, mesNasc, anoNasc] = dataNascimento.split('/').map(Number);
        const nascimento = new Date(anoNasc, mesNasc - 1, diaNasc);
        const diff = new Date() - nascimento;
        const idade = Math.floor(diff / (1000 * 60 * 60 * 24 * 365.25));
        return idade;
    };

    const idade = data.dataNascimento ? calcularIdade(data.dataNascimento) : 'N/A';

    doc.autoTable({
        startY: 70,
        head: [
            ["ITEM", "NOME", "DATA DO \nNASCIMENTO", "IDADE", "BAIRRO", "MUNICÍPIO", "DATA DO \nFALECIMENTO", "LOCAL DO FALECIMENTO", "CAUSA DA MORTE", "CEMITÉRIO"]
        ],
        body: [
            [
                "1",
                data.nome || "Não informado",
                data.dataNascimento || "Não informado",
                idade,
                data.bairro || "Não informado", // Adicione um campo "Bairro" se necessário
                data.municipio || "Teresina", // Padrão para Teresina
                data.dataObito || "Não informado",
                data.localObito || "Não informado",
                data.causaMorte || "Não informado",
                data.cemiterio || "Não informado"
            ]
        ],
        styles: {
            fontSize: 8, // Ajustado para caber melhor
            valign: 'middle',
            halign: 'center', // Centraliza o texto horizontalmente
            cellPadding: 1, // Reduz o padding para economizar espaço
            overflow: 'linebreak' // Garante que o texto quebre a linha se necessário
        },
        columnStyles: {
            0: { cellWidth: 15 }, // ITEM
            1: { cellWidth: 30 }, // NOME
            2: { cellWidth: 30 }, // DATA DO NASCIMENTO
            3: { cellWidth: 15 }, // IDADE
            4: { cellWidth: 25 }, // BAIRRO
            5: { cellWidth: 25 }, // MUNICÍPIO
            6: { cellWidth: 30 }, // DATA DO FALECIMENTO
            7: { cellWidth: 30 }, // LOCAL DO FALECIMENTO
            8: { cellWidth: 25 }, // CAUSA DA MORTE
            9: { cellWidth: 30 }  // CEMITÉRIO
        },
        headStyles: {
            fillColor: [240, 240, 240], // Cor de fundo do cabeçalho (teal)
            textColor: [0, 0, 0], // Cor do texto do cabeçalho (branco)
            fontSize: 10, // Tamanho da fonte do cabeçalho
            fontStyle: 'bold',
            halign: 'center', // Centraliza o texto do cabeçalho
            valign: 'middle'
        },
        theme: 'grid',
        margin: { left: 10, right: 10 } // Margens para evitar corte
    });

    doc.save(`Relatorio ${dia}_${mes}_${ano}.pdf`);
}


async function gerarTabelaDatmPDF(dataSEMF) {
    const { jsPDF } = window.jspdf;
    const anoAtual = new Date().getFullYear();

    const doc = new jsPDF({
        orientation: "landscape",
        unit: "mm",
        format: "a4",
    });

    const logoUrl = "https://i.imgur.com/vaDmtrQ.png";
    const logoWidth = 25;
    const logoHeight = logoWidth * 1318 / 1200;
    const logoX = 10;
    const logoY = 10;
    const imageBase64 = await loadImageAsBase64(logoUrl);
    if (imageBase64) {
        doc.addImage(imageBase64, 'PNG', logoX, logoY, logoWidth, logoHeight);
    }

    const distancia_logo_texto = 5;
    const texto_cabecalhoX = logoX + logoWidth + distancia_logo_texto;
    const texto_cabecalhoY = logoY + 5;

    doc.setFontSize(11);
    doc.setFont(undefined, 'bold');
    doc.text("ESTADO DO PIAUÍ", texto_cabecalhoX, texto_cabecalhoY);
    doc.setFont(undefined, 'normal');
    doc.text("Prefeitura Municipal de Teresina", texto_cabecalhoX, texto_cabecalhoY + 6);
    doc.text("Superintendência de Desenvolvimento Urbano - Centro-Norte", texto_cabecalhoX, texto_cabecalhoY + 12);
    doc.text("Plantão Funerário", texto_cabecalhoX, texto_cabecalhoY + 18);

    const titulo = `TABELA DE PREÇOS DATM ${anoAtual}`;
    doc.setFontSize(14);
    doc.setFont(undefined, 'bold');
    const pageWidth = doc.internal.pageSize.getWidth();
    const textWidth = doc.getTextWidth(titulo);
    const centerX = (pageWidth - textWidth) / 2;
    const titleY = 45;
    doc.text(titulo, centerX, titleY); // Removido o sublinhado

    doc.autoTable({
        startY: 55,
        head: [
            [
                { content: "CEMITÉRIOS DO GRUPO A\n- SÃO JOSÉ (Norte);\n- SÃO JUDAS TADEU (Leste);", colSpan: 4 },
                { content: "CEMITÉRIOS DO GRUPO B\n- DOM BOSCO (Sul);\n- SANTO ANTONIO (Norte);\n- RENASCENÇA (Sudeste)", colSpan: 4 },
                { content: "CEMITÉRIOS DO GRUPO C\n- POTY VELHO (Norte);\n- SÃO SEBASTIÃO (Sudeste);\n- SÃO JOÃO BATISTA (Norte);\n- MORROS (Leste);\n- SANTA CRUZ (Sul);\n- SANTA MARIA (Norte);\n- SANTA MONICA (Leste);", colSpan: 4 }
            ]
        ],
        body: [
            [
                "TIPO", "CÓDIGO", "VALOR (R$)", "TOTAL (R$)",
                "TIPO", "CÓDIGO", "VALOR (R$)", "TOTAL (R$)",
                "TIPO", "CÓDIGO", "VALOR (R$)", "TOTAL (R$)"
            ],
            [
                "ADULTO 1ª ABERTURA/COVA RASA", "14020102\n14020104", "0,00\n0,00", "0,00",
                "ADULTO 1ª ABERTURA/COVA RASA", "14020202\n14020204", "44,45\n44,45", "88,90",
                "ADULTO 1ª ABERTURA/COVA RASA", "14020302\n14020304", "17,28\n17,28", "34,56"
            ],
            [
                "ADULTO JAZIGO/GAVETA", "14020103\n14020105", "112,37\n58,04", "170,41",
                "ADULTO JAZIGO/GAVETA", "14020203\n14020205", "85,17\n44,45", "129,62",
                "ADULTO JAZIGO/GAVETA", "14020303\n14020305", "30,87\n17,28", "48,15"
            ],
            [
                "INFANTE 1ª ABERTURA/COVA RASA", "14020402\n14020404", "0,00\n0,00", "0,00",
                "INFANTE 1ª ABERTURA/COVA RASA", "14020502\n14020504", "0,00\n0,00", "0,00",
                "INFANTE 1ª ABERTURA/COVA RASA", "14020602\n14020604", "10,53\n10,53", "21,06"
            ],
            [
                "INFANTE JAZIGO/GAVETA", "14020403\n14020405", "81,45\n54,32", "135,77",
                "INFANTE JAZIGO/GAVETA", "14020503\n14020505", "0,00\n0,00", "0,00",
                "INFANTE JAZIGO/GAVETA", "14020603\n14020605", "24,03\n24,03", "48,06"
            ]
        ],
        styles: {
            fontSize: 8,
            valign: 'middle',
            halign: 'left',
            cellPadding: 1,
            overflow: 'linebreak'
        },
        columnStyles: {
            0: { cellWidth: 30, halign: 'left'},  // TIPO (Grupo A)
            1: { cellWidth: 25, halign: 'center'},  // CÓDIGO (Grupo A)
            2: { cellWidth: 15, halign: 'center' },  // VALOR (Grupo A)
            3: { cellWidth: 15, halign: 'center' },  // TOTAL (Grupo A)
            4: { cellWidth: 30, halign: 'left' },  // TIPO (Grupo B)
            5: { cellWidth: 25, halign: 'center' },  // CÓDIGO (Grupo B)
            6: { cellWidth: 15, halign: 'center' },  // VALOR (Grupo B)
            7: { cellWidth: 15, halign: 'center' },  // TOTAL (Grupo B)
            8: { cellWidth: 30, halign: 'left' },  // TIPO (Grupo C)
            9: { cellWidth: 25, halign: 'center' },  // CÓDIGO (Grupo C)
            10: { cellWidth: 15, halign: 'center' }, // VALOR (Grupo C)
            11: { cellWidth: 15, halign: 'center' }  // TOTAL (Grupo C)
        },
        headStyles: {
            fillColor: [240, 240, 240],
            textColor: [0, 0, 0],
            fontSize: 8,
            fontStyle: 'bold',
            halign: 'left',
            valign: 'middle'
        },
        theme: 'grid',
        margin: { left: 10, right: 10 }
    });

    doc.save(`Tabela_de_Precos_DATM_${anoAtual}.pdf`);
}




export {loadImageAsBase64, exportJsonToExcel, readExcelFile, mergeExcelData, gerarGuiaDeSepultamentoPDF, gerarRelatorioFmsPDF, gerarRelatorioSduPDF, gerarTabelaDatmPDF}