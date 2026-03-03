/**
 * FuzzyMatching.gs
 * Fuzzy string matching utilities for K&L Recycling CRM
 * Implements Levenshtein distance, Jaro-Winkler, and normalized similarity
 */

const FuzzyMatching = {
  /**
   * Calculate Levenshtein distance between two strings
   * @param {string} a - First string
   * @param {string} b - Second string
   * @returns {number} - Edit distance (0 = identical)
   */
  levenshteinDistance: function(a, b) {
    if (!a || !b) return Math.max(a ? a.length : 0, b ? b.length : 0);
    
    const matrix = [];
    const aLen = a.length;
    const bLen = b.length;
    
    // Initialize first column
    for (let i = 0; i <= aLen; i++) {
      matrix[i] = [i];
    }
    
    // Initialize first row
    for (let j = 0; j <= bLen; j++) {
      matrix[0][j] = j;
    }
    
    // Fill matrix
    for (let i = 1; i <= aLen; i++) {
      for (let j = 1; j <= bLen; j++) {
        const cost = a[i - 1] === b[j - 1] ? 0 : 1;
        matrix[i][j] = Math.min(
          matrix[i - 1][j] + 1,      // deletion
          matrix[i][j - 1] + 1,      // insertion
          matrix[i - 1][j - 1] + cost // substitution
        );
      }
    }
    
    return matrix[aLen][bLen];
  },

  /**
   * Calculate similarity score (0-1) based on Levenshtein distance
   * @param {string} a - First string
   * @param {string} b - Second string
   * @returns {number} - Similarity score (1 = identical, 0 = completely different)
   */
  levenshteinSimilarity: function(a, b) {
    if (!a && !b) return 1;
    if (!a || !b) return 0;
    
    const distance = this.levenshteinDistance(a, b);
    const maxLength = Math.max(a.length, b.length);
    return maxLength === 0 ? 1 : 1 - (distance / maxLength);
  },

  /**
   * Calculate Jaro similarity
   * @param {string} a - First string
   * @param {string} b - Second string
   * @returns {number} - Jaro similarity (0-1)
   */
  jaroSimilarity: function(a, b) {
    if (!a || !b) return 0;
    if (a === b) return 1;
    
    const aLen = a.length;
    const bLen = b.length;
    
    const matchDistance = Math.floor(Math.max(aLen, bLen) / 2) - 1;
    const aMatches = new Array(aLen).fill(false);
    const bMatches = new Array(bLen).fill(false);
    
    let matches = 0;
    let transpositions = 0;
    
    // Find matches
    for (let i = 0; i < aLen; i++) {
      const start = Math.max(0, i - matchDistance);
      const end = Math.min(i + matchDistance + 1, bLen);
      
      for (let j = start; j < end; j++) {
        if (bMatches[j] || a[i] !== b[j]) continue;
        aMatches[i] = true;
        bMatches[j] = true;
        matches++;
        break;
      }
    }
    
    if (matches === 0) return 0;
    
    // Count transpositions
    let k = 0;
    for (let i = 0; i < aLen; i++) {
      if (!aMatches[i]) continue;
      while (!bMatches[k]) k++;
      if (a[i] !== b[k]) transpositions++;
      k++;
    }
    
    return ((matches / aLen) + (matches / bLen) + ((matches - transpositions / 2) / matches)) / 3;
  },

  /**
   * Calculate Jaro-Winkler similarity (better for short strings like names)
   * @param {string} a - First string
   * @param {string} b - Second string
   * @returns {number} - Jaro-Winkler similarity (0-1)
   */
  jaroWinklerSimilarity: function(a, b) {
    if (!a || !b) return 0;
    if (a === b) return 1;
    
    const jaro = this.jaroSimilarity(a, b);
    
    // Find common prefix length (max 4)
    let prefixLen = 0;
    const maxPrefix = Math.min(4, a.length, b.length);
    while (prefixLen < maxPrefix && a[prefixLen] === b[prefixLen]) {
      prefixLen++;
    }
    
    // Winkler boost
    const boost = 0.1;
    return jaro + (prefixLen * boost * (1 - jaro));
  },

  /**
   * Normalize string for comparison
   * Removes special chars, extra spaces, converts to lowercase
   * @param {string} str - Input string
   * @returns {string} - Normalized string
   */
  normalize: function(str) {
    if (!str) return '';
    return str
      .toString()
      .toLowerCase()
      .replace(/[,.]/g, '')
      .replace(/\b(llc|inc|corp|co|ltd|company)\b/g, '')
      .replace(/[^a-z0-9\s]/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  },

  /**
   * Calculate fuzzy similarity using multiple algorithms
   * @param {string} a - First string
   * @param {string} b - Second string
   * @returns {number} - Combined similarity score (0-1)
   */
  similarity: function(a, b) {
    if (!a || !b) return 0;
    if (a === b) return 1;
    
    const normalizedA = this.normalize(a);
    const normalizedB = this.normalize(b);
    
    if (normalizedA === normalizedB) return 1;
    
    // Combine multiple similarity measures
    const levenshtein = this.levenshteinSimilarity(normalizedA, normalizedB);
    const jaroWinkler = this.jaroWinklerSimilarity(normalizedA, normalizedB);
    
    // Weight Jaro-Winkler higher (better for names)
    return (levenshtein * 0.4) + (jaroWinkler * 0.6);
  },

  /**
   * Check if two strings are fuzzy matches (above threshold)
   * @param {string} a - First string
   * @param {string} b - Second string
   * @param {number} threshold - Minimum similarity (0-1), default 0.85
   * @returns {boolean} - True if fuzzy match
   */
  isMatch: function(a, b, threshold = 0.85) {
    return this.similarity(a, b) >= threshold;
  },

  /**
   * Find best fuzzy match for a search term in a list
   * @param {string} searchTerm - Term to search for
   * @param {Array<string>} candidates - List of candidate strings
   * @param {number} threshold - Minimum similarity threshold
   * @returns {Object|null} - Best match with score, or null
   */
  findBestMatch: function(searchTerm, candidates, threshold = 0.7) {
    if (!searchTerm || !candidates || candidates.length === 0) return null;
    
    const normalizedSearch = this.normalize(searchTerm);
    let bestMatch = null;
    let bestScore = 0;
    
    candidates.forEach(candidate => {
      const score = this.similarity(normalizedSearch, this.normalize(candidate));
      if (score > bestScore && score >= threshold) {
        bestScore = score;
        bestMatch = candidate;
      }
    });
    
    return bestMatch ? { match: bestMatch, score: bestScore } : null;
  },

  /**
   * Fuzzy search - returns all matches above threshold sorted by score
   * @param {string} searchTerm - Term to search for
   * @param {Array<Object>} items - Array of objects to search
   * @param {string} key - Object key to search in
   * @param {number} threshold - Minimum similarity (default 0.6)
   * @returns {Array<Object>} - Matches with similarity scores
   */
  fuzzySearch: function(searchTerm, items, key, threshold = 0.6) {
    if (!searchTerm || !items || items.length === 0) return [];
    
    const normalizedSearch = this.normalize(searchTerm);
    const results = [];
    
    items.forEach(item => {
      const value = item[key];
      if (!value) return;
      
      const score = this.similarity(normalizedSearch, this.normalize(value));
      if (score >= threshold) {
        results.push({
          item: item,
          score: score,
          similarity: Math.round(score * 100) + '%'
        });
      }
    });
    
    // Sort by score descending
    return results.sort((a, b) => b.score - a.score);
  },

  /**
   * Find potential duplicates using fuzzy matching
   * @param {Array<Object>} records - Array of company records
   * @param {string} key - Key to compare (e.g., 'Company Name')
   * @param {number} threshold - Similarity threshold (default 0.9)
   * @returns {Array<Object>} - Array of potential duplicate groups
   */
  findPotentialDuplicates: function(records, key, threshold = 0.9) {
    if (!records || records.length < 2) return [];
    
    const duplicates = [];
    const processed = new Set();
    
    for (let i = 0; i < records.length; i++) {
      if (processed.has(i)) continue;
      
      const recordA = records[i];
      const nameA = recordA[key];
      if (!nameA) continue;
      
      const group = [recordA];
      
      for (let j = i + 1; j < records.length; j++) {
        if (processed.has(j)) continue;
        
        const recordB = records[j];
        const nameB = recordB[key];
        if (!nameB) continue;
        
        const similarity = this.similarity(nameA, nameB);
        
        if (similarity >= threshold) {
          group.push(recordB);
          processed.add(j);
        }
      }
      
      if (group.length > 1) {
        duplicates.push({
          canonicalName: nameA,
          similarity: Math.round(this.averageSimilarity(group, key) * 100) + '%',
          records: group
        });
      }
      
      processed.add(i);
    }
    
    return duplicates;
  },

  /**
   * Calculate average pairwise similarity within a group
   * @private
   */
  averageSimilarity: function(group, key) {
    let total = 0;
    let count = 0;
    
    for (let i = 0; i < group.length; i++) {
      for (let j = i + 1; j < group.length; j++) {
        total += this.similarity(group[i][key], group[j][key]);
        count++;
      }
    }
    
    return count === 0 ? 0 : total / count;
  },

  /**
   * Soundex algorithm for phonetic matching
   * @param {string} str - Input string
   * @returns {string} - 4-character Soundex code
   */
  soundex: function(str) {
    if (!str) return '';
    
    str = str.toUpperCase().replace(/[^A-Z]/g, '');
    if (str.length === 0) return '';
    
    const soundexMap = {
      'B': '1', 'F': '1', 'P': '1', 'V': '1',
      'C': '2', 'G': '2', 'J': '2', 'K': '2', 'Q': '2', 'S': '2', 'X': '2', 'Z': '2',
      'D': '3', 'T': '3',
      'L': '4',
      'M': '5', 'N': '5',
      'R': '6'
    };
    
    let result = str[0];
    let prevCode = soundexMap[str[0]];
    
    for (let i = 1; i < str.length && result.length < 4; i++) {
      const code = soundexMap[str[i]];
      if (code && code !== prevCode) {
        result += code;
        prevCode = code;
      }
    }
    
    return result.padEnd(4, '0');
  },

  /**
   * Check if two strings sound alike (phonetic match)
   * @param {string} a - First string
   * @param {string} b - Second string
   * @returns {boolean} - True if they have the same Soundex code
   */
  soundsAlike: function(a, b) {
    return this.soundex(a) === this.soundex(b);
  }
};

/**
 * Wrapper function for fuzzy search in sidebar
 * @param {string} term - Search term
 * @returns {Array<Object>} - Matching companies with scores
 */
function searchCompaniesFuzzy(term) {
  if (!term || term.length < 2) return [];
  
  const allData = DataLayer.getMultipleRecords(['PROSPECTS', 'ACTIVE_CONTAINERS']);
  const prospects = allData.PROSPECTS || [];
  const active = allData.ACTIVE_CONTAINERS || [];
  
  const allCompanies = [...prospects, ...active];
  
  // Fuzzy search on company name
  const results = FuzzyMatching.fuzzySearch(term, allCompanies, 'Company Name', 0.6);
  
  // Also do exact ID matching
  const exactIdMatches = allCompanies.filter(c => 
    c['Company ID']?.toLowerCase().includes(term.toLowerCase())
  ).map(c => ({ item: c, score: 1, similarity: '100%' }));
  
  // Combine and deduplicate
  const seen = new Set();
  const combined = [...exactIdMatches, ...results].filter(r => {
    const id = r.item['Company ID'];
    if (seen.has(id)) return false;
    seen.add(id);
    return true;
  });
  
  // Return top 10
  return combined.slice(0, 10).map(r => ({
    id: r.item['Company ID'],
    name: r.item['Company Name'],
    city: r.item.City || 'N/A',
    score: r.similarity,
    type: r.item._type || (prospects.includes(r.item) ? 'Prospect' : 'Active')
  }));
}

/**
 * Enhanced duplicate detection with fuzzy matching
 * @returns {Array<Object>} - Groups of potential duplicates
 */
function findFuzzyDuplicates() {
  try {
    const allData = DataLayer.getMultipleRecords(['PROSPECTS', 'ACTIVE_CONTAINERS']);
    const allRecords = [
      ...(allData.PROSPECTS || []),
      ...(allData.ACTIVE_CONTAINERS || [])
    ];
    
    // Find duplicates with 0.85 similarity threshold
    const duplicates = FuzzyMatching.findPotentialDuplicates(
      allRecords, 
      'Company Name', 
      0.85
    );
    
    // Format for display
    return duplicates.map(group => ({
      name: group.canonicalName,
      similarity: group.similarity,
      count: group.records.length,
      ids: group.records.map(r => r['Company ID']),
      records: group.records
    }));
    
  } catch (error) {
    Logger.log('Error in findFuzzyDuplicates: ' + error.message);
    return [];
  }
}

/**
 * Check if a new company name might be a duplicate
 * @param {string} newName - New company name to check
 * @returns {Array<Object>} - Potential existing matches
 */
function checkForSimilarCompany(newName) {
  if (!newName) return [];
  
  const allData = DataLayer.getMultipleRecords(['PROSPECTS', 'ACTIVE_CONTAINERS']);
  const allRecords = [
    ...(allData.PROSPECTS || []),
    ...(allData.ACTIVE_CONTAINERS || [])
  ];
  
  const matches = [];
  const normalizedNew = FuzzyMatching.normalize(newName);
  
  allRecords.forEach(record => {
    const existingName = record['Company Name'];
    if (!existingName) return;
    
    const similarity = FuzzyMatching.similarity(normalizedNew, existingName);
    
    if (similarity >= 0.8) {
      matches.push({
        companyId: record['Company ID'],
        companyName: existingName,
        similarity: Math.round(similarity * 100) + '%',
        score: similarity
      });
    }
  });
  
  return matches.sort((a, b) => b.score - a.score);
}
