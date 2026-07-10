var fixedColumnLetters = {
  country: "P",
  major: "V",
  businessLine: ["AN", "AZ", "BL"],
};

function normalizeMatrix(matrix) {
  const headerIndex = detectHeaderRow(matrix);
  const seenHeaders = new Map();
  const headerLine = matrix[headerIndex] || [];
  const headers = Array.from({ length: headerLine.length }, (_, index) => {
    const value = headerLine[index];
    const label = clean(value);
    return uniqueHeader(label || `Column ${index + 1}`, seenHeaders);
  });
  const rows = matrix.slice(headerIndex + 1).map((line) => {
    const row = {};
    headers.forEach((header, index) => {
      row[header] = normalizeCell(line[index]);
    });
    return row;
  }).filter((row) => Object.values(row).some((value) => clean(value)));
  return { headers, rows };
}

function uniqueHeader(label, seenHeaders) {
  const count = seenHeaders.get(label) || 0;
  seenHeaders.set(label, count + 1);
  return count ? `${label}.${count}` : label;
}

function buildPerson(row, headers, key) {
  const get = (field) => row[findColumn(headers, columnAliases[field])] || "";
  const getFixed = (letter) => row[columnFromLetter(headers, letter)] || "";
  const person = {
    key,
    email: get("email"),
    name: get("name"),
    country: getFixed(fixedColumnLetters.country) || get("country"),
    city: get("city"),
    pm: get("pm"),
    degree: get("degree"),
    major: getFixed(fixedColumnLetters.major) || get("major"),
    domain: get("domain"),
    english: get("english"),
    hours: get("hours"),
    yearsExperience: get("yearsExperience"),
    projectCount: get("projectCount"),
    languageCountry: get("languageCountry"),
    nativeLanguage: get("nativeLanguage"),
    foreignLanguage: get("foreignLanguage"),
    completedAt: get("completedAt"),
    sourceRows: [row],
    projects: [],
  };
  mergeProjectExperience(person, row, headers);
  return person;
}

function mergeProjectExperience(person, row, headers) {
  if (!person.sourceRows.includes(row)) person.sourceRows.push(row);
  for (let group = 0; group < 3; group += 1) {
    const suffix = group === 0 ? "" : `.${group}`;
    const project = {};
    projectColumnBases.forEach((base) => {
      const column = headers.find((header) => header === `${base}${suffix}`);
      if (column) project[base] = row[column] || "";
    });
    const fixedBusinessLine = row[columnFromLetter(headers, fixedColumnLetters.businessLine[group])] || "";
    if (fixedBusinessLine) project["Business Line"] = fixedBusinessLine;
    if (Object.values(project).some((value) => clean(value))) {
      person.projects.push(project);
    }
  }
}

function columnFromLetter(headers, letter) {
  const index = columnLetterToIndex(letter);
  return headers[index] || "";
}

function columnLetterToIndex(letter) {
  return clean(letter)
    .toUpperCase()
    .split("")
    .reduce((index, char) => index * 26 + char.charCodeAt(0) - 64, 0) - 1;
}
