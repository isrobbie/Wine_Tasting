// GOOGLE APPS SCRIPT - Presentation Generator
// Add this to your Apps Script project (same one with the form handler)
// Run the generatePresentation() function after collecting responses

function generatePresentation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  const headers = data[0];
  const responses = data.slice(1);
  
  if (responses.length === 0) {
    SpreadsheetApp.getUi().alert('No responses found!');
    return;
  }
  
  const stats = calculateStats(responses);
  
  const presentation = SlidesApp.create('Wine Tasting Results - ' + new Date().toLocaleDateString());
  const slides = presentation.getSlides();
  if (slides.length > 0) slides[0].remove();
  
  createTitleSlide(presentation);
  createWinnersSlide(presentation, stats);
  createRankingsSlide(presentation, stats);
  createTasterAwardsSlide(presentation, stats, responses);
  createThankYouSlide(presentation);
  
  const url = presentation.getUrl();
  SpreadsheetApp.getUi().alert('Success!', 'Presentation created:\n' + url, SpreadsheetApp.getUi().ButtonSet.OK);
}

function calculateStats(responses) {
  const wines = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'];
  const categories = {
    'A': 'Sparkling', 'B': 'Sparkling', 'C': 'Sparkling',
    'D': 'White', 'E': 'White', 'F': 'White',
    'G': 'Red', 'H': 'Red', 'I': 'Red'
  };
  const wineNames = {
    'A': 'Langlois-Chateau Crémant de Loire Brut',
    'B': 'Bouvet Ladubay Saumur Brut',
    'C': 'Maison Antech Blanquette de Limoux Brut',
    'D': 'Krasno Sauvignon Blanc-Ribolla Gialla',
    'E': 'Forrest Wines \'The Doctors\' Sauvignon Blanc',
    'F': 'Beau Rivage Bordeaux Blanc',
    'G': 'Porta 6',
    'H': 'LB7 Red',
    'I': 'La Belle Angèle Pinot Noir'
  };
  
  const wineStats = {};
  
  wines.forEach((wine, idx) => {
    const ratingCol = (idx * 2) + 2;
    const ratings = responses
      .map(row => parseFloat(row[ratingCol]))
      .filter(r => r > 0 && !isNaN(r));
    
    if (ratings.length > 0) {
      const avg = ratings.reduce((a, b) => a + b, 0) / ratings.length;
      const min = Math.min(...ratings);
      const max = Math.max(...ratings);
      const variance = ratings.reduce((sum, r) => sum + Math.pow(r - avg, 2), 0) / ratings.length;
      const stdDev = Math.sqrt(variance);
      
      wineStats[wine] = {
        letter: wine,
        name: wineNames[wine],
        category: categories[wine],
        average: avg,
        min: min,
        max: max,
        stdDev: stdDev,
        count: ratings.length,
        ratings: ratings,
        hasRatings: true
      };
    } else {
      wineStats[wine] = {
        letter: wine,
        name: wineNames[wine],
        category: categories[wine],
        average: 0,
        min: 0,
        max: 0,
        stdDev: 0,
        count: 0,
        ratings: [],
        hasRatings: false
      };
    }
  });
  
  const sorted = Object.keys(wineStats).sort((a, b) => {
    if (!wineStats[a].hasRatings) return 1;
    if (!wineStats[b].hasRatings) return -1;
    return wineStats[b].average - wineStats[a].average;
  });
  
  const categoryWinners = {};
  ['Sparkling', 'White', 'Red'].forEach(cat => {
    const catWines = Object.keys(wineStats).filter(w => wineStats[w].category === cat && wineStats[w].hasRatings);
    if (catWines.length > 0) {
      categoryWinners[cat] = catWines.sort((a, b) => wineStats[b].average - wineStats[a].average)[0];
    }
  });
  
  return {
    wines: wineStats,
    sorted: sorted,
    categoryWinners: categoryWinners
  };
}

function createTitleSlide(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill('#6D2E46');
  
  const title = slide.insertTextBox('WINE TASTING');
  title.setLeft(100).setTop(150).setWidth(500).setHeight(80);
  const t1 = title.getText();
  t1.getTextStyle().setFontSize(60).setFontFamily('Arial').setBold(true).setForegroundColor('#FFFFFF');
  t1.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  
  const subtitle = slide.insertTextBox('RESULTS REVEALED');
  subtitle.setLeft(100).setTop(250).setWidth(500).setHeight(60);
  const t2 = subtitle.getText();
  t2.getTextStyle().setFontSize(44).setFontFamily('Arial').setBold(true).setForegroundColor('#ECE2D0');
  t2.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

function createWinnersSlide(presentation, stats) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill('#FFFFFF');
  
  const title = slide.insertTextBox('THE WINNERS');
  title.setLeft(50).setTop(30).setWidth(600).setHeight(60);
  const titleText = title.getText();
  titleText.getTextStyle().setFontSize(48).setFontFamily('Arial').setBold(true).setForegroundColor('#6D2E46');
  titleText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  
  const categories = [
    { name: 'Best Sparkling', key: 'Sparkling', y: 110 },
    { name: 'Best White', key: 'White', y: 210 },
    { name: 'Best Red', key: 'Red', y: 310 }
  ];
  
  categories.forEach(cat => {
    const catLabel = slide.insertTextBox(cat.name);
    catLabel.setLeft(120).setTop(cat.y).setWidth(200).setHeight(30);
    const catText = catLabel.getText();
    catText.getTextStyle().setFontSize(18).setFontFamily('Arial').setBold(true).setForegroundColor('#6D2E46');
    
    if (stats.categoryWinners[cat.key]) {
      const winner = stats.wines[stats.categoryWinners[cat.key]];
      
      const wineLetter = slide.insertTextBox('Wine ' + winner.letter);
      wineLetter.setLeft(120).setTop(cat.y + 32).setWidth(150).setHeight(30);
      const wineText = wineLetter.getText();
      wineText.getTextStyle().setFontSize(24).setFontFamily('Arial').setBold(true).setForegroundColor('#4CAF50');
      
      const wineName = slide.insertTextBox(winner.name);
      wineName.setLeft(290).setTop(cat.y + 32).setWidth(280).setHeight(30);
      const nameText = wineName.getText();
      nameText.getTextStyle().setFontSize(10).setFontFamily('Arial').setForegroundColor('#666666');
      
      const rating = slide.insertTextBox('Rating: ' + winner.average.toFixed(1));
      rating.setLeft(120).setTop(cat.y + 62).setWidth(150).setHeight(22);
      const ratingText = rating.getText();
      ratingText.getTextStyle().setFontSize(16).setFontFamily('Arial').setForegroundColor('#666666');
    } else {
      const noRating = slide.insertTextBox('n/a');
      noRating.setLeft(120).setTop(cat.y + 32).setWidth(150).setHeight(30);
      const naText = noRating.getText();
      naText.getTextStyle().setFontSize(24).setFontFamily('Arial').setBold(true).setForegroundColor('#999999');
    }
  });
}

function createRankingsSlide(presentation, stats) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill('#ECE2D0');
  
  const title = slide.insertTextBox('FULL RANKINGS');
  title.setLeft(50).setTop(20).setWidth(600).setHeight(50);
  const titleText = title.getText();
  titleText.getTextStyle().setFontSize(40).setFontFamily('Arial').setBold(true).setForegroundColor('#6D2E46');
  titleText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  
  // Table headers
  const headers = ['Rank', 'Wine', 'Category', 'Avg', 'Range', 'Std Dev'];
  const colWidths = [40, 250, 80, 60, 70, 70];
  const colStarts = [50, 90, 340, 420, 480, 550];
  
  let yPos = 85;
  headers.forEach((header, i) => {
    const box = slide.insertTextBox(header);
    box.setLeft(colStarts[i]).setTop(yPos).setWidth(colWidths[i]).setHeight(25);
    const text = box.getText();
    text.getTextStyle().setFontSize(12).setFontFamily('Arial').setBold(true).setForegroundColor('#3D1B2C');
  });
  
  yPos = 115;
  let actualRank = 1;
  stats.sorted.forEach((wineLetter) => {
    const wine = stats.wines[wineLetter];
    const rowData = [
      wine.hasRatings ? actualRank.toString() : '-',
      'Wine ' + wine.letter + ' - ' + wine.name,
      wine.category,
      wine.hasRatings ? wine.average.toFixed(1) : 'n/a',
      wine.hasRatings ? wine.min + '-' + wine.max : 'n/a',
      wine.hasRatings ? wine.stdDev.toFixed(1) : 'n/a'
    ];
    
    rowData.forEach((data, i) => {
      const box = slide.insertTextBox(data);
      box.setLeft(colStarts[i]).setTop(yPos).setWidth(colWidths[i]).setHeight(22);
      const text = box.getText();
      const fontSize = i === 1 ? 8 : 11; // Smaller font for wine name column
      text.getTextStyle().setFontSize(fontSize).setFontFamily('Arial').setForegroundColor('#333333');
      if (actualRank === 1 && wine.hasRatings) {
        text.getTextStyle().setBold(true).setForegroundColor('#4CAF50');
      }
    });
    
    if (wine.hasRatings) actualRank++;
    yPos += 28;
  });
}

function createTasterAwardsSlide(presentation, stats, responses) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill('#6D2E46');
  
  const title = slide.insertTextBox('TASTER AWARDS');
  title.setLeft(50).setTop(25).setWidth(600).setHeight(50);
  const titleText = title.getText();
  titleText.getTextStyle().setFontSize(44).setFontFamily('Arial').setBold(true).setForegroundColor('#FFFFFF');
  titleText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  
  const tasterStats = [];
  responses.forEach(row => {
    const name = row[1];
    const ratings = [];
    for (let i = 2; i < row.length; i += 2) {
      const rating = parseFloat(row[i]);
      if (rating > 0 && !isNaN(rating)) {
        ratings.push(rating);
      }
    }
    if (ratings.length > 0) {
      const avg = ratings.reduce((a, b) => a + b, 0) / ratings.length;
      const variance = ratings.reduce((sum, r) => sum + Math.pow(r - avg, 2), 0) / ratings.length;
      const stdDev = Math.sqrt(variance);
      tasterStats.push({ name, avg, stdDev, ratings });
    }
  });
  
  if (tasterStats.length === 0) return;
  
  // Calculate correlations for Most/Least Correlated
  let maxCorr = -2, minCorr = 2;
  let maxPair = ['', ''], minPair = ['', ''];
  
  for (let i = 0; i < tasterStats.length; i++) {
    for (let j = i + 1; j < tasterStats.length; j++) {
      const corr = calculateCorrelation(tasterStats[i].ratings, tasterStats[j].ratings, stats);
      if (corr > maxCorr) {
        maxCorr = corr;
        maxPair = [tasterStats[i].name, tasterStats[j].name];
      }
      if (corr < minCorr) {
        minCorr = corr;
        minPair = [tasterStats[i].name, tasterStats[j].name];
      }
    }
  }
  
  const highest = tasterStats.sort((a, b) => b.avg - a.avg)[0];
  const lowest = tasterStats.sort((a, b) => a.avg - b.avg)[0];
  const mostUnique = tasterStats.sort((a, b) => b.stdDev - a.stdDev)[0];
  const mostNormal = tasterStats.sort((a, b) => a.stdDev - b.stdDev)[0];
  
  const awards = [
    { label: 'Highest Rater:', value: highest.name, y: 90 },
    { label: 'Lowest Rater:', value: lowest.name, y: 145 },
    { label: 'Most Unique:', value: mostUnique.name, y: 200 },
    { label: 'Most Normal:', value: mostNormal.name, y: 255 },
    { label: 'Most Correlated:', value: maxPair[0] + ' & ' + maxPair[1], y: 310 },
    { label: 'Least Correlated:', value: minPair[0] + ' & ' + minPair[1], y: 365 }
  ];
  
  awards.forEach(award => {
    const label = slide.insertTextBox(award.label);
    label.setLeft(100).setTop(award.y).setWidth(500).setHeight(25);
    const labelText = label.getText();
    labelText.getTextStyle().setFontSize(16).setFontFamily('Arial').setBold(true).setForegroundColor('#ECE2D0');
    
    const value = slide.insertTextBox(award.value);
    value.setLeft(100).setTop(award.y + 25).setWidth(500).setHeight(25);
    const valueText = value.getText();
    valueText.getTextStyle().setFontSize(18).setFontFamily('Arial').setBold(true).setForegroundColor('#FFFFFF');
  });
}

function calculateCorrelation(ratings1, ratings2, stats) {
  const wines = Object.keys(stats.wines);
  const pairs = [];
  
  for (let i = 0; i < Math.min(ratings1.length, ratings2.length); i++) {
    if (ratings1[i] > 0 && ratings2[i] > 0) {
      pairs.push([ratings1[i], ratings2[i]]);
    }
  }
  
  if (pairs.length < 2) return 0;
  
  const mean1 = pairs.reduce((sum, p) => sum + p[0], 0) / pairs.length;
  const mean2 = pairs.reduce((sum, p) => sum + p[1], 0) / pairs.length;
  
  let num = 0, den1 = 0, den2 = 0;
  pairs.forEach(p => {
    const diff1 = p[0] - mean1;
    const diff2 = p[1] - mean2;
    num += diff1 * diff2;
    den1 += diff1 * diff1;
    den2 += diff2 * diff2;
  });
  
  if (den1 === 0 || den2 === 0) return 0;
  return num / Math.sqrt(den1 * den2);
}

function createThankYouSlide(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill('#A26769');
  
  const title = slide.insertTextBox('CHEERS!');
  title.setLeft(50).setTop(180).setWidth(600).setHeight(80);
  const titleText = title.getText();
  titleText.getTextStyle().setFontSize(72).setFontFamily('Arial').setBold(true).setForegroundColor('#FFFFFF');
  titleText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  
  const subtitle = slide.insertTextBox('Thanks for sharing your palates with us');
  subtitle.setLeft(50).setTop(280).setWidth(600).setHeight(40);
  const subtitleText = subtitle.getText();
  subtitleText.getTextStyle().setFontSize(20).setFontFamily('Arial').setItalic(true).setForegroundColor('#ECE2D0');
  subtitleText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}
