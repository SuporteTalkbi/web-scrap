const express = require('express');
const axios = require('axios');
const cheerio = require('cheerio');
const app = express();

app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
  next();
});

app.get('/scrape', async (req, res) => {
  try {
    // Simulando navegador com headers
    const response = await axios.get('https://1drv.ms/x/c/80098ff3d93ac023/EaCXh5CRkLpPm1z84N8zndoBFQI8vTry8s1cASNezDq-NQ?e=L1oXmv', {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Accept': 'text/html,application/xhtml+xml',
        'Accept-Language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7'
      }
    });
    
    // Usar Cheerio para extrair dados
    const $ = cheerio.load(response.data);
    const rows = [];
    
    // Extrair dados da tabela
    $('table tr').each((i, row) => {
      const rowData = [];
      $(row).find('td, th').each((j, cell) => {
        rowData.push($(cell).text().trim());
      });
      if (rowData.some(text => text !== '')) {
        rows.push(rowData);
      }
    });
    
    // Formatar dados
    if (rows.length > 0) {
      const headers = rows[0];
      const result = rows.slice(1).map(row => {
        const item = {};
        headers.forEach((header, index) => {
          if (header) {
            item[header] = index < row.length ? row[index] : '';
          }
        });
        return item;
      });
      res.json(result);
    } else {
      res.status(404).json({error: 'Nenhum dado encontrado'});
    }
  } catch (error) {
    res.status(500).json({error: error.message});
  }
});

app.get('/', (req, res) => {
  res.send('Excel Scraper está funcionando! Use /scrape para obter dados da planilha.');
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Servidor rodando na porta ${port}`);
});
