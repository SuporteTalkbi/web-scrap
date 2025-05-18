const express = require('express');
const axios = require('axios');
const cheerio = require('cheerio');
const app = express();

app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
  next();
});

// Função para extrair dados da tabela HTML
function extrairDadosTabela(html) {
  const $ = cheerio.load(html);
  const rows = [];
  
  // Tenta encontrar qualquer tabela na página
  const tabelas = $('table');
  console.log(`Encontradas ${tabelas.length} tabelas na página`);
  
  // Se nenhuma tabela for encontrada, tenta buscar divs que parecem tabelas
  if (tabelas.length === 0) {
    console.log('Buscando estruturas de tabela alternativas...');
    // Procura por estruturas que possam ser tabelas (Excel Online usa divs para render)
    $('.excelGrid, .excelTable, .sheetContainer, [role="grid"]').each((i, element) => {
      console.log(`Encontrada possível estrutura de tabela ${i}`);
      // Código para extrair dados de estruturas div que simulam tabelas
      // (Isso é específico para a estrutura do Excel Online)
    });
  }
  
  // Processa cada tabela encontrada (geralmente a primeira é a que queremos)
  tabelas.first().find('tr').each((i, row) => {
    const rowData = [];
    $(row).find('td, th').each((j, cell) => {
      rowData.push($(cell).text().trim());
    });
    
    if (rowData.some(text => text !== '')) {
      rows.push(rowData);
    }
  });
  
  return rows;
}

// Tenta processar os dados da tabela para formato JSON
function processarDados(rows) {
  if (!rows || rows.length === 0) {
    return [];
  }
  
  // Se não encontrar dados em formato de tabela, log para debug
  console.log(`Encontradas ${rows.length} linhas de dados`);
  
  // Assume a primeira linha como cabeçalho
  const headers = rows[0];
  console.log('Cabeçalhos encontrados:', headers);
  
  // Converte o resto para objetos
  return rows.slice(1).map(row => {
    const item = {};
    headers.forEach((header, index) => {
      if (header && header.trim()) {
        item[header] = index < row.length ? row[index] : '';
      }
    });
    return item;
  });
}

// Rota principal para scraping
app.get('/scrape', async (req, res) => {
  try {
    console.log('Iniciando scraping da planilha...');
    
    // Tenta primeiro com AllOrigins (geralmente o mais confiável)
    try {
      console.log('Tentando com AllOrigins...');
      const url = 'https://1drv.ms/x/c/80098ff3d93ac023/EaCXh5CRkLpPm1z84N8zndoBFQI8vTry8s1cASNezDq-NQ?e=L1oXmv';
      const response = await axios.get('https://api.allorigins.win/raw?url=' + encodeURIComponent(url), {
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        },
        timeout: 30000
      });
      
      console.log('Resposta recebida, analisando HTML...');
      const rows = extrairDadosTabela(response.data);
      const result = processarDados(rows);
      
      if (result.length > 0) {
        console.log(`Sucesso! Encontrados ${result.length} registros.`);
        return res.json(result);
      }
      
      console.log('Nenhum dado encontrado com AllOrigins, tentando alternativa...');
    } catch (error) {
      console.log('Erro ao usar AllOrigins:', error.message);
    }
    
    // Segunda tentativa com CORS Proxy IO
    try {
      console.log('Tentando com CORS Proxy IO...');
      const url = 'https://1drv.ms/x/c/80098ff3d93ac023/EaCXh5CRkLpPm1z84N8zndoBFQI8vTry8s1cASNezDq-NQ?e=L1oXmv';
      const response = await axios.get('https://corsproxy.io/?' + encodeURIComponent(url), {
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        },
        timeout: 30000
      });
      
      console.log('Resposta recebida de CORS Proxy IO, analisando HTML...');
      const rows = extrairDadosTabela(response.data);
      const result = processarDados(rows);
      
      if (result.length > 0) {
        console.log(`Sucesso! Encontrados ${result.length} registros.`);
        return res.json(result);
      }
      
      console.log('Nenhum dado encontrado com CORS Proxy IO');
    } catch (error) {
      console.log('Erro ao usar CORS Proxy IO:', error.message);
    }
    
    // Terceira tentativa com acesso direto (pode falhar devido a CORS)
    try {
      console.log('Tentando acesso direto...');
      const url = 'https://1drv.ms/x/c/80098ff3d93ac023/EaCXh5CRkLpPm1z84N8zndoBFQI8vTry8s1cASNezDq-NQ?e=L1oXmv';
      const response = await axios.get(url, {
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
          'Accept': 'text/html,application/xhtml+xml,application/xml',
          'Accept-Language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7'
        },
        timeout: 30000
      });
      
      console.log('Resposta direta recebida, analisando HTML...');
      const rows = extrairDadosTabela(response.data);
      const result = processarDados(rows);
      
      if (result.length > 0) {
        console.log(`Sucesso! Encontrados ${result.length} registros.`);
        return res.json(result);
      }
    } catch (error) {
      console.log('Erro no acesso direto:', error.message);
    }
    
    // Se todas as tentativas falharem
    console.log('Todas as tentativas falharam');
    res.status(404).json({
      error: 'Não foi possível extrair dados da tabela',
      message: 'A estrutura da página do Excel Online pode ter mudado ou está protegida contra scraping.',
      possíveis_soluções: [
        'Tente fazer download da planilha e usar API oficial do Excel/OneDrive',
        'Verifique se a URL da planilha não requer autenticação'
      ]
    });
    
  } catch (error) {
    console.error('Erro geral:', error);
    res.status(500).json({ error: error.message });
  }
});

// Rota de diagnóstico
app.get('/', (req, res) => {
  res.send('Excel Scraper está funcionando! Use /scrape para obter dados da planilha.');
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Servidor rodando na porta ${port}`);
});
