const express = require('express');
const puppeteer = require('puppeteer');
const app = express();

// Adicionar CORS para permitir requisições do n8n
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
  next();
});

// Rota principal para scraping
app.get('/scrape', async (req, res) => {
  console.log('Iniciando scraping...');
  
  // Iniciar o navegador
  const browser = await puppeteer.launch({
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
    executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || '/usr/bin/chromium-browser'
  });
  const page = await browser.newPage();
  
  try {
    // Configurar navegador para parecer usuário real
    await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36');
    await page.setViewport({ width: 1920, height: 1080 });
    
    console.log('Acessando URL do Excel...');
    await page.goto(process.env.SHEET_URL || 'process.env.SHEET_URL', {
      waitUntil: 'networkidle2',
      timeout: 90000
    });
    
    console.log('Aguardando carregamento da tabela...');
    // Tenta vários seletores possíveis para aumentar as chances de sucesso
    const seletorEncontrado = await Promise.race([
      page.waitForSelector('table', { timeout: 60000 }).then(() => 'table'),
      page.waitForSelector('.excelTable', { timeout: 60000 }).then(() => '.excelTable'),
      page.waitForSelector('.excelGrid', { timeout: 60000 }).then(() => '.excelGrid'),
      page.waitForSelector('.excel-table', { timeout: 60000 }).then(() => '.excel-table')
    ]).catch(() => null);
    
    if (!seletorEncontrado) {
      console.log('Nenhum seletor encontrado. Tentando localizar elementos da tabela via DOM...');
    } else {
      console.log(`Seletor encontrado: ${seletorEncontrado}`);
    }
    
    // Aguarda tempo adicional para garantir carregamento completo
    await page.waitForTimeout(5000);
    
    // Tira screenshot para debug (útil para ver o que o scraper está vendo)
    await page.screenshot({ path: '/tmp/excel-page.png' });
    console.log('Screenshot salvo');
    
    // Extrair dados
    console.log('Extraindo dados da tabela...');
    const data = await page.evaluate(() => {
      // Função para encontrar a tabela, tentando vários seletores
      function encontrarTabela() {
        // Tenta seletores específicos primeiro
        const seletores = ['table', '.excelTable', '.excelGrid', '.excel-table'];
        for (const seletor of seletores) {
          const elements = document.querySelectorAll(seletor);
          if (elements.length > 0) return elements[0];
        }
        
        // Se não encontrar, procura elementos que parecem uma tabela
        // Procura div com muitas divs filhas em sequência (provável linha)
        const divs = Array.from(document.querySelectorAll('div'));
        const possiveisTables = divs.filter(div => {
          const children = Array.from(div.children);
          return children.length > 3 && 
                 children.every(child => child.tagName === children[0].tagName);
        });
        
        if (possiveisTables.length > 0) {
          return possiveisTables[0]; // retorna o primeiro candidato
        }
        
        return null;
      }
      
      // Função para extrair linhas e células
      function extrairDados(elemento) {
        if (!elemento) return [];
        
        // Verifica se é uma tabela HTML padrão
        if (elemento.tagName === 'TABLE') {
          return Array.from(elemento.querySelectorAll('tr')).map(row => {
            return Array.from(row.querySelectorAll('td, th')).map(cell => 
              cell.textContent.trim()
            );
          });
        }
        
        // Para divs estruturadas como tabela
        const childRows = Array.from(elemento.children);
        return childRows.map(row => {
          // Se for uma div que contém células
          if (row.children.length > 0) {
            return Array.from(row.children).map(cell => 
              cell.textContent.trim()
            );
          }
          return [row.textContent.trim()];
        });
      }
      
      const tabela = encontrarTabela();
      if (!tabela) return [];
      
      const dados = extrairDados(tabela);
      return dados.filter(row => row.some(cell => cell !== ''));
    });
    
    console.log(`Dados extraídos: ${data.length} linhas`);
    
    if (data.length === 0) {
      console.log('Nenhum dado encontrado. Verificando HTML da página...');
      const html = await page.content();
      console.log(`Tamanho do HTML: ${html.length} caracteres`);
      
      res.status(404).json({ 
        error: 'Não foi possível extrair dados da tabela',
        message: 'Verifique se a URL do Excel Online é acessível sem login'
      });
      return;
    }
    
    // Processar dados
    console.log('Processando dados extraídos...');
    const headers = data[0];
    const result = data.slice(1).map(row => {
      const item = {};
      headers.forEach((header, index) => {
        if (header) { // Evita colunas sem cabeçalho
          item[header] = index < row.length ? row[index] : '';
        }
      });
      return item;
    });
    
    console.log('Enviando resposta com dados processados');
    res.json(result);
  } catch (error) {
    console.error('Erro durante scraping:', error);
    res.status(500).json({ 
      error: error.message,
      stack: error.stack
    });
  } finally {
    await browser.close();
    console.log('Navegador fechado');
  }
});

// Rota de diagnóstico
app.get('/', (req, res) => {
  res.send('Excel Scraper está funcionando! Use /scrape para obter dados da planilha.');
});

// Iniciar servidor
const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Servidor rodando na porta ${port}`);
});
