import React, { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import { AlertCircle, BookOpen, Building, Calculator, Scale, TrendingUp, Users } from 'lucide-react';

const MUNICIPALITY_OPTIONS = [
  { value: 'do1999', label: 'Do 1 999 obyvatel', landCoefficient: 1.2, datasetLabel: 'Do 2 000' },
  { value: '2000-9999', label: '2 000-9 999 obyvatel', landCoefficient: 1.25, datasetLabel: '2 000 - 9 999' },
  { value: '10000-49999', label: '10 000-49 999 obyvatel', landCoefficient: 1.3, datasetLabel: '10 000 - 49 999' },
  { value: 'nad50000', label: '50 000 a vice obyvatel', landCoefficient: 1.35, datasetLabel: '50 000' }
];

const PERSON_OPTIONS = [
  { value: '0', label: '0 (zadna zohlednovana osoba)' },
  { value: '1', label: '1 osoba' },
  { value: '2-3', label: '2 az 3 osoby' },
  { value: '4+', label: '4 a vice osob' }
];

const QUARTERS = [
  { value: '1', label: '1. ctvrtleti' },
  { value: '2', label: '2. ctvrtleti' },
  { value: '3', label: '3. ctvrtleti' },
  { value: '4', label: '4. ctvrtleti' }
];

const getTodayIso = () => new Date().toISOString().slice(0, 10);
const QUARTER_LABELS = Object.fromEntries(QUARTERS.map((quarter) => [quarter.value, quarter.label]));

const normalizeText = (value) =>
  String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9]+/g, ' ')
    .trim()
    .toLowerCase();

const parseNumber = (value) => {
  if (value === '' || value === null || value === undefined) return null;
  const parsed = Number(String(value).replace(',', '.').trim());
  return Number.isFinite(parsed) ? parsed : null;
};

const formatCurrency = (value) =>
  value === null || value === undefined || Number.isNaN(value)
    ? '-'
    : `${new Intl.NumberFormat('cs-CZ', { minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(value)} Kc`;

const formatNumber = (value, digits = 2) =>
  value === null || value === undefined || Number.isNaN(value)
    ? '-'
    : new Intl.NumberFormat('cs-CZ', { minimumFractionDigits: digits, maximumFractionDigits: digits }).format(value);

const roundToTwo = (value) => Math.round(value * 100) / 100;
const parseIsoDate = (value) => {
  if (!value) return null;
  const date = new Date(`${value}T00:00:00`);
  return Number.isNaN(date.getTime()) ? null : date;
};
const getQuarterFromDate = (date) => String(Math.floor(date.getMonth() / 3) + 1);
const parseQuarterKey = (key) => {
  const match = /^(\d{4})-Q([1-4])$/.exec(key);
  return match ? { year: Number(match[1]), quarter: Number(match[2]) } : null;
};
const formatQuarterKeyLabel = (key) => {
  const parsed = parseQuarterKey(key);
  return parsed ? `${QUARTER_LABELS[String(parsed.quarter)]} ${parsed.year}` : '-';
};
const getRequiredPricePeriod = (decisionDateValue) => {
  const decisionDate = parseIsoDate(decisionDateValue);
  if (!decisionDate) return null;
  const decisionYear = decisionDate.getFullYear();
  return { startYear: decisionYear - 4, endYear: decisionYear - 2 };
};
const getEligibleCurrentQuarterKey = (indices, decisionDateValue) => {
  const decisionDate = parseIsoDate(decisionDateValue);
  if (!decisionDate) return null;

  const decisionYear = decisionDate.getFullYear();
  const decisionQuarter = Number(getQuarterFromDate(decisionDate));
  const eligibleKeys = Object.keys(indices).filter((key) => {
    const parsed = parseQuarterKey(key);
    if (!parsed) return false;
    return parsed.year < decisionYear || (parsed.year === decisionYear && parsed.quarter <= decisionQuarter);
  });

  if (eligibleKeys.length > 0) {
    return eligibleKeys.sort((left, right) => {
      const a = parseQuarterKey(left);
      const b = parseQuarterKey(right);
      return a.year - b.year || a.quarter - b.quarter;
    }).at(-1);
  }

  return Object.keys(indices)
    .sort((left, right) => {
      const a = parseQuarterKey(left);
      const b = parseQuarterKey(right);
      return a.year - b.year || a.quarter - b.quarter;
    })
    .at(-1) ?? null;
};
const getComparableHistoricQuarterKey = (currentQuarterKey) => {
  const parsed = parseQuarterKey(currentQuarterKey);
  return parsed ? `${parsed.year - 3}-Q${parsed.quarter}` : null;
};
const getMunicipalityDatasetKey = (value) => {
  const normalized = normalizeText(value);
  if (normalized.includes('do 2 000') || normalized.includes('mene nez 2 000')) return 'do 2 000';
  if (normalized.includes('2 000 az 10 000') || normalized.includes('vice nebo rovno 2 000')) return '2 000 - 9 999';
  if (normalized.includes('10 000 az 50 000') || normalized.includes('vice nebo rovno 10 000')) return '10 000 - 49 999';
  if (normalized.includes('od 50 000') || normalized.includes('vice nebo rovno 50 000')) return '50 000';
  return normalized;
};

const getLandCoefficient = (municipalitySize) =>
  MUNICIPALITY_OPTIONS.find((item) => item.value === municipalitySize)?.landCoefficient ?? 1;

const getPersonCoefficient = (type, persons) => {
  if (type === 'byt') {
    return { 0: 0.85, 1: 1, '2-3': 1.3, '4+': 1.6 }[persons] ?? null;
  }
  return { 0: 0.7, 1: 1, '2-3': 1, '4+': 1.15 }[persons] ?? null;
};

const parseXml = (text) => new DOMParser().parseFromString(text, 'application/xml');
const childText = (node, tag) => node.querySelector(`:scope > ${tag}`)?.textContent?.trim() ?? '';
const childTags = (node, tag) =>
  Array.from(node.children)
    .filter((child) => child.tagName === tag)
    .map((child) => child.textContent.trim());

function parsePriceDataset(xmlText) {
  const doc = parseXml(xmlText);
  const vecById = Object.fromEntries(
    Array.from(doc.querySelectorAll('metaSlovnik > vecneUpresneni > element')).map((element) => [
      element.getAttribute('ID'),
      childText(element, 'text')
    ])
  );

  const territories = Array.from(doc.querySelectorAll('metaSlovnik > uzemi > element')).map((element) => ({
    id: element.getAttribute('ID'),
    text: childText(element, 'text'),
    ciselnik: childText(element, 'ciselnik')
  }));

  const territoryById = Object.fromEntries(territories.map((territory) => [territory.id, territory]));
  const periodNode = doc.querySelector('metaSlovnik > obdobi > cas');
  const period = periodNode ? { from: childText(periodNode, 'casOd'), to: childText(periodNode, 'casDo') } : null;
  const districtOptions = territories
    .filter((territory) => territory.ciselnik === '101')
    .map((territory) => territory.text)
    .sort((a, b) => a.localeCompare(b, 'cs'));

  const districtPrices = {};
  const nationalPrices = {};

  Array.from(doc.querySelectorAll('data > udaj')).forEach((item) => {
    const value = parseNumber(childText(item, 'hod'));
    const territory = territoryById[childText(item, 'uze')];
    const vecTexts = childTags(item, 'vec').map((id) => vecById[id] ?? '');
    const typeText = vecTexts.find((text) => text.includes('Byt') || text.includes('Rodinn'));
    const municipalityText = vecTexts.find((text) => text.includes('2 000') || text.includes('10 000') || text.includes('50 000') || text.includes('Od 50'));

    if (!territory || value === null || !typeText) return;

    const typeKey = typeText.includes('Byt') ? 'byt' : 'dum';

    if (normalizeText(territory.text) === 'ceska republika') {
      if (!municipalityText) nationalPrices[typeKey] = value;
      return;
    }

    if (territory.ciselnik !== '101' || !municipalityText) return;

    const districtKey = normalizeText(territory.text);
    districtPrices[districtKey] ??= { label: territory.text, byt: {}, dum: {} };
    districtPrices[districtKey][typeKey][getMunicipalityDatasetKey(municipalityText)] = value;
  });

  return { districtPrices, nationalPrices, districtOptions, period };
}

function parseIndexDataset(xmlText) {
  const doc = parseXml(xmlText);
  const vecById = Object.fromEntries(
    Array.from(doc.querySelectorAll('metaSlovnik > vecneUpresneni > element')).map((element) => [
      element.getAttribute('ID'),
      childText(element, 'text')
    ])
  );

  const territoryById = Object.fromEntries(
    Array.from(doc.querySelectorAll('metaSlovnik > uzemi > element')).map((element) => [
      element.getAttribute('ID'),
      childText(element, 'text')
    ])
  );

  const periodById = Object.fromEntries(
    Array.from(doc.querySelectorAll('metaSlovnik > obdobi > cas')).map((element) => [
      element.getAttribute('ID'),
      childText(element, 'casOd')
    ])
  );

  const result = {};

  Array.from(doc.querySelectorAll('data > udaj')).forEach((item) => {
    const territory = territoryById[childText(item, 'uze')];
    if (normalizeText(territory) !== 'ceska republika') return;

    const vecTexts = childTags(item, 'vec').map((id) => vecById[id] ?? '');
    const matchesTargetSeries =
      vecTexts.some((text) => text.includes('Byt')) &&
      vecTexts.some((text) => normalizeText(text).includes('prumer')) &&
      vecTexts.some((text) => normalizeText(text).includes('podil')) &&
      vecTexts.some((text) => normalizeText(text).includes('index'));

    if (!matchesTargetSeries) return;

    const value = parseNumber(childText(item, 'hod'));
    const from = periodById[childText(item, 'cas')];
    if (value === null || !from) return;

    const date = new Date(from);
    result[`${date.getFullYear()}-Q${Math.floor(date.getMonth() / 3) + 1}`] = value;
  });

  return result;
}

function parseHouseSizeWorkbook(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });

  const districtAreas = {};
  const nationalAreas = {};
  const regionalAreas = {};
  let currentRegion = '';
  let period = null;

  rows.forEach((row) => {
    const rowText = row.map((cell) => String(cell ?? '')).join(' ');
    if (!period) {
      const match = rowText.match(/(20\d{2})\s*-\s*(20\d{2})/);
      if (match) {
        period = { startYear: Number(match[1]), endYear: Number(match[2]) };
      }
    }

    const region = String(row[1] ?? '').trim();
    const district = String(row[2] ?? '').trim();
    const avgArea = parseNumber(row[3]);
    const avgPrice = parseNumber(row[4]);

    if (region && !region.startsWith('Celkem ')) {
      currentRegion = region;
    }

    if (avgArea === null) return;

    const normalizedRegion = normalizeText(region);
    const normalizedDistrict = normalizeText(district);

    if (normalizedRegion === 'celkem cr') {
      nationalAreas.dum = avgArea;
      return;
    }

    if (region.startsWith('Celkem ')) {
      regionalAreas[normalizeText(region.replace('Celkem ', '').trim())] = avgArea;
      return;
    }

    if (district && avgPrice !== null) {
      districtAreas[normalizedDistrict] = { dum: avgArea, region: currentRegion || region };
      return;
    }

    if (normalizedRegion === 'hlavni mesto praha' && !district) {
      regionalAreas[normalizedRegion] = avgArea;
    }
  });

  return { districtAreas, nationalAreas, regionalAreas, period };
}

function parseApartmentSizeWorkbook(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });

  const districtAreas = {};
  const nationalAreas = {};
  const regionalAreas = {};
  let currentRegion = '';
  let period = null;

  rows.forEach((row) => {
    const rowText = row.map((cell) => String(cell ?? '')).join(' ');
    if (!period) {
      const match = rowText.match(/(20\d{2})\s*-\s*(20\d{2})/);
      if (match) {
        period = { startYear: Number(match[1]), endYear: Number(match[2]) };
      }
    }

    const region = String(row[0] ?? '').trim();
    const district = String(row[1] ?? '').trim();
    const avgArea = parseNumber(row[2]);
    const avgPrice = parseNumber(row[3]);

    if (region && !region.startsWith('Celkem ')) {
      currentRegion = region;
    }

    if (avgArea === null) return;

    const normalizedRegion = normalizeText(region);
    const normalizedDistrict = normalizeText(district);

    if (normalizedRegion === 'celkem cr') {
      nationalAreas.byt = avgArea;
      return;
    }

    if (region.startsWith('Celkem ')) {
      regionalAreas[normalizeText(region.replace('Celkem ', '').trim())] = avgArea;
      return;
    }

    if (district && avgPrice !== null) {
      districtAreas[normalizedDistrict] = { byt: avgArea, region: currentRegion || region };
      return;
    }

    if (normalizedRegion === 'hlavni mesto praha' && district === 'Praha') {
      regionalAreas[normalizedRegion] = avgArea;
    }
  });

  return { districtAreas, nationalAreas, regionalAreas, period };
}

function parseQuarterIndexWorkbook(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });

  const yearsRow = rows[0] ?? [];
  const quarterRow = rows[1] ?? [];
  const nationalRow = rows.find((row) => normalizeText(row[0]) === 'cr');
  const indices = {};

  if (!nationalRow) {
    return indices;
  }

  let currentYear = '';
  for (let columnIndex = 2; columnIndex < nationalRow.length; columnIndex += 1) {
    const yearCell = String(yearsRow[columnIndex] ?? '').trim();
    if (/^\d{4}$/.test(yearCell)) {
      currentYear = yearCell;
    }

    const quarterLabel = String(quarterRow[columnIndex] ?? '').trim();
    const value = parseNumber(nationalRow[columnIndex]);
    const quarterMatch = /ctv\s*([1-4])|ctv([1-4])|ctvrtleti\s*([1-4])|čtv\s*([1-4])|čtv([1-4])/i.exec(
      normalizeText(quarterLabel).replace('ctv', 'ctv ')
    );
    const quarter =
      quarterMatch?.[1] ?? quarterMatch?.[2] ?? quarterMatch?.[3] ?? quarterMatch?.[4] ?? quarterMatch?.[5] ?? null;

    if (/^\d{4}$/.test(currentYear) && quarter && value !== null) {
      indices[`${currentYear}-Q${quarter}`] = value;
    }
  }

  return indices;
}

function App() {
  const [type, setType] = useState('dum');
  const [municipalitySize, setMunicipalitySize] = useState('do1999');
  const [districtLabel, setDistrictLabel] = useState('');
  const [regionLabel, setRegionLabel] = useState('');
  const [decisionDate, setDecisionDate] = useState(getTodayIso());
  const [districtAreaInput, setDistrictAreaInput] = useState('');
  const [nationalAreaInput, setNationalAreaInput] = useState('');
  const [persons, setPersons] = useState('2-3');
  const [datasets, setDatasets] = useState({ loading: true, error: '', prices: null, indices: {} });
  const [houseSizeWorkbook, setHouseSizeWorkbook] = useState({
    loaded: false,
    fileName: '0140162502.xlsx',
    error: '',
    districtAreas: {},
    nationalAreas: {},
    regionalAreas: {},
    period: null
  });
  const [apartmentSizeWorkbook, setApartmentSizeWorkbook] = useState({
    loaded: false,
    fileName: '0140162507.xlsx',
    error: '',
    districtAreas: {},
    nationalAreas: {},
    regionalAreas: {},
    period: null
  });

  useEffect(() => {
    let cancelled = false;

    (async () => {
      try {
        const [priceText, index2024Text, index2023Text, apartmentIndexWorkbookResponse] = await Promise.all([
          fetch('/CEN13A.xml').then((response) => response.text()),
          fetch('/CEN15A1.xml').then((response) => response.text()),
          fetch('/CEN15A1%20(2023).xml').then((response) => response.text()),
          fetch('/0140162508.xlsx').then((response) => response.arrayBuffer())
        ]);

        if (cancelled) return;

        const prices = parsePriceDataset(priceText);
        const indices = {
          ...parseQuarterIndexWorkbook(apartmentIndexWorkbookResponse),
          ...parseIndexDataset(index2023Text),
          ...parseIndexDataset(index2024Text)
        };
        setDatasets({ loading: false, error: '', prices, indices });
      } catch {
        if (!cancelled) {
          setDatasets({ loading: false, error: 'Nepodarilo se nacist lokalni XML soubory CSU.', prices: null, indices: {} });
        }
      }
    })();

    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    let cancelled = false;

    (async () => {
      try {
        const response = await fetch('/0140162507.xlsx');
        const buffer = await response.arrayBuffer();
        if (cancelled) return;
        const parsed = parseApartmentSizeWorkbook(buffer);
        setApartmentSizeWorkbook({
          loaded: true,
          fileName: '0140162507.xlsx',
          error: '',
          districtAreas: parsed.districtAreas,
          nationalAreas: parsed.nationalAreas,
          regionalAreas: parsed.regionalAreas,
          period: parsed.period
        });
      } catch {
        if (!cancelled) {
          setApartmentSizeWorkbook((prev) => ({
            ...prev,
            loaded: false,
            error: 'Soubor 0140162507.xlsx se nepodarilo nacist automaticky.'
          }));
        }
      }
    })();

    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    let cancelled = false;

    (async () => {
      try {
        const response = await fetch('/0140162502.xlsx');
        const buffer = await response.arrayBuffer();
        if (cancelled) return;
        const parsed = parseHouseSizeWorkbook(buffer);
        setHouseSizeWorkbook({
          loaded: true,
          fileName: '0140162502.xlsx',
          error: '',
          districtAreas: parsed.districtAreas,
          nationalAreas: parsed.nationalAreas,
          regionalAreas: parsed.regionalAreas,
          period: parsed.period
        });
      } catch {
        if (!cancelled) {
          setHouseSizeWorkbook((prev) => ({
            ...prev,
            loaded: false,
            error: 'Soubor 0140162502.xlsx se nepodarilo nacist automaticky.'
          }));
        }
      }
    })();

    return () => {
      cancelled = true;
    };
  }, []);

  const municipalityKey = useMemo(
    () => getMunicipalityDatasetKey(MUNICIPALITY_OPTIONS.find((item) => item.value === municipalitySize)?.datasetLabel ?? ''),
    [municipalitySize]
  );

  const districtMatch = useMemo(
    () => datasets.prices?.districtPrices?.[normalizeText(districtLabel)] ?? null,
    [datasets.prices, districtLabel]
  );

  const autoDistrictPrice = useMemo(
    () => districtMatch?.[type]?.[municipalityKey] ?? null,
    [districtMatch, municipalityKey, type]
  );

  const autoNationalPrice = useMemo(
    () => datasets.prices?.nationalPrices?.[type] ?? null,
    [datasets.prices, type]
  );

  const autoDistrictArea = useMemo(
    () =>
      type === 'dum'
        ? houseSizeWorkbook.districtAreas?.[normalizeText(districtLabel)]?.dum ?? null
        : apartmentSizeWorkbook.districtAreas?.[normalizeText(districtLabel)]?.byt ?? null,
    [apartmentSizeWorkbook, houseSizeWorkbook, districtLabel, type]
  );

  useEffect(() => {
    const detectedRegion =
      (type === 'dum'
        ? houseSizeWorkbook.districtAreas?.[normalizeText(districtLabel)]?.region
        : apartmentSizeWorkbook.districtAreas?.[normalizeText(districtLabel)]?.region) ?? '';
    if (detectedRegion) {
      setRegionLabel(detectedRegion);
    }
  }, [apartmentSizeWorkbook, houseSizeWorkbook, districtLabel, type]);

  const regionalFallbackArea = useMemo(
    () =>
      type === 'dum'
        ? houseSizeWorkbook.regionalAreas?.[normalizeText(regionLabel)] ?? null
        : apartmentSizeWorkbook.regionalAreas?.[normalizeText(regionLabel)] ?? null,
    [apartmentSizeWorkbook, houseSizeWorkbook, regionLabel, type]
  );

  const autoNationalArea = useMemo(
    () => (type === 'dum' ? houseSizeWorkbook.nationalAreas?.dum ?? null : apartmentSizeWorkbook.nationalAreas?.byt ?? null),
    [apartmentSizeWorkbook, houseSizeWorkbook, type]
  );

  const requiredPricePeriod = useMemo(
    () => getRequiredPricePeriod(decisionDate),
    [decisionDate]
  );

  const loadedPricePeriod = useMemo(() => {
    const fromDate = parseIsoDate(datasets.prices?.period?.from);
    const toDate = parseIsoDate(datasets.prices?.period?.to);
    return fromDate && toDate ? { startYear: fromDate.getFullYear(), endYear: toDate.getFullYear() } : null;
  }, [datasets.prices]);

  const currentQuarterKey = useMemo(
    () => getEligibleCurrentQuarterKey(datasets.indices, decisionDate),
    [datasets.indices, decisionDate]
  );

  const historicQuarterKey = useMemo(
    () => getComparableHistoricQuarterKey(currentQuarterKey),
    [currentQuarterKey]
  );

  const currentIndex = useMemo(
    () => (currentQuarterKey ? datasets.indices[currentQuarterKey] ?? null : null),
    [datasets.indices, currentQuarterKey]
  );

  const historicIndex = useMemo(
    () => (historicQuarterKey ? datasets.indices[historicQuarterKey] ?? null : null),
    [datasets.indices, historicQuarterKey]
  );

  const calculation = useMemo(() => {
    const districtArea = autoDistrictArea ?? regionalFallbackArea ?? parseNumber(districtAreaInput);
    const nationalArea = autoNationalArea ?? parseNumber(nationalAreaInput);
    const kp = type === 'dum' ? getLandCoefficient(municipalitySize) : 1;
    const ko = getPersonCoefficient(type, persons);
    const raw = districtArea !== null && autoDistrictPrice !== null ? districtArea * autoDistrictPrice : null;
    const nationalRef = nationalArea !== null && autoNationalPrice !== null ? nationalArea * autoNationalPrice : null;
    const min = nationalRef !== null ? nationalRef * 0.2 : null;
    const max = nationalRef !== null ? nationalRef * 2 : null;
    const bounded = raw === null || min === null || max === null ? raw : Math.min(Math.max(raw, min), max);
    const boundsApplied = raw !== null && bounded !== raw;
    const growth = currentIndex !== null && historicIndex !== null && historicIndex !== 0 ? roundToTwo(currentIndex / historicIndex) : null;
    const protectedValue =
      bounded !== null && growth !== null && ko !== null ? Math.ceil(bounded * kp * growth * ko) : null;

    const missing = [];
    if (!districtLabel) missing.push('okres');
    if (districtArea === null) missing.push(type === 'dum' ? 'okresni nebo krajova velikost domu' : 'okresni velikost bytu');
    if (autoDistrictPrice === null) missing.push('automaticky nalezena kupni cena v okrese');
    if (nationalArea === null) missing.push('republikova prumerna velikost');
    if (autoNationalPrice === null) missing.push('automaticky nalezena republikova kupni cena');
    if (
      requiredPricePeriod &&
      loadedPricePeriod &&
      (requiredPricePeriod.startYear !== loadedPricePeriod.startYear || requiredPricePeriod.endYear !== loadedPricePeriod.endYear)
    ) {
      missing.push(`cenovy soubor pro obdobi ${requiredPricePeriod.startYear}-${requiredPricePeriod.endYear}`);
    }
    if (
      type === 'dum' &&
      requiredPricePeriod &&
      houseSizeWorkbook.period &&
      (requiredPricePeriod.startYear !== houseSizeWorkbook.period.startYear || requiredPricePeriod.endYear !== houseSizeWorkbook.period.endYear)
    ) {
      missing.push(`soubor velikosti domu pro obdobi ${requiredPricePeriod.startYear}-${requiredPricePeriod.endYear}`);
    }
    if (
      type === 'byt' &&
      requiredPricePeriod &&
      apartmentSizeWorkbook.period &&
      (requiredPricePeriod.startYear !== apartmentSizeWorkbook.period.startYear || requiredPricePeriod.endYear !== apartmentSizeWorkbook.period.endYear)
    ) {
      missing.push(`soubor velikosti bytu pro obdobi ${requiredPricePeriod.startYear}-${requiredPricePeriod.endYear}`);
    }
    if (currentIndex === null) missing.push(`aktualni index pro ${formatQuarterKeyLabel(currentQuarterKey)}`);
    if (historicIndex === null) {
      missing.push(`historicky index pro ${formatQuarterKeyLabel(historicQuarterKey)} (chybi starsi ctvrtletni soubor CSU)`);
    }

    return { districtArea, nationalArea, kp, ko, raw, min, max, bounded, boundsApplied, growth, protectedValue, missing };
  }, [
    autoDistrictArea,
    autoDistrictPrice,
    autoNationalArea,
    autoNationalPrice,
    apartmentSizeWorkbook.period,
    currentIndex,
    districtAreaInput,
    districtLabel,
    historicIndex,
    houseSizeWorkbook.period,
    loadedPricePeriod,
    municipalitySize,
    nationalAreaInput,
    persons,
    requiredPricePeriod,
    regionalFallbackArea,
    type
  ]);

  const currentQuarterLabel = formatQuarterKeyLabel(currentQuarterKey);
  const historicQuarterLabel = formatQuarterKeyLabel(historicQuarterKey);
  const requiredPricePeriodLabel = requiredPricePeriod ? `${requiredPricePeriod.startYear}-${requiredPricePeriod.endYear}` : '-';
  const loadedPricePeriodLabel = loadedPricePeriod ? `${loadedPricePeriod.startYear}-${loadedPricePeriod.endYear}` : '-';
  const loadedHousePeriodLabel = houseSizeWorkbook.period ? `${houseSizeWorkbook.period.startYear}-${houseSizeWorkbook.period.endYear}` : '-';
  const loadedApartmentPeriodLabel = apartmentSizeWorkbook.period ? `${apartmentSizeWorkbook.period.startYear}-${apartmentSizeWorkbook.period.endYear}` : '-';
  const periodsAligned = Boolean(
    requiredPricePeriod &&
      loadedPricePeriod &&
      houseSizeWorkbook.period &&
      apartmentSizeWorkbook.period &&
      requiredPricePeriod.startYear === loadedPricePeriod.startYear &&
      requiredPricePeriod.endYear === loadedPricePeriod.endYear &&
      requiredPricePeriod.startYear === houseSizeWorkbook.period.startYear &&
      requiredPricePeriod.endYear === houseSizeWorkbook.period.endYear &&
      requiredPricePeriod.startYear === apartmentSizeWorkbook.period.startYear &&
      requiredPricePeriod.endYear === apartmentSizeWorkbook.period.endYear
  );

  return (
    <div className="min-h-screen bg-[#efe3d3] text-slate-800 p-4 md:p-8 font-sans">
      <div className="max-w-6xl mx-auto space-y-6">
        <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
          <div className="flex items-center gap-4 mb-2">
            <div className="p-3 bg-blue-100 text-blue-700 rounded-xl">
              <Calculator size={28} />
            </div>
            <div>
              <h1 className="text-2xl font-bold text-slate-900">Vypocet chraneneho obydli v insolvenci</h1>
            </div>
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
          <div className="lg:col-span-7 space-y-6">
            <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
              <h2 className="text-lg font-semibold border-b pb-4 mb-5 flex items-center gap-2">
                <Building size={20} className="text-slate-500" />
                Rozhodne skutecnosti a zdroje
              </h2>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
                <div className="space-y-2">
                  <label className="block text-sm font-medium text-slate-700">Typ obydli</label>
                  <select value={type} onChange={(e) => setType(e.target.value)} className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg">
                    <option value="byt">Byt</option>
                    <option value="dum">Rodinny dum</option>
                  </select>
                </div>

                <div className="space-y-2">
                  <label className="block text-sm font-medium text-slate-700">Datum rozhodneho okamziku</label>
                  <input type="date" value={decisionDate} onChange={(e) => setDecisionDate(e.target.value)} className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg" />
                </div>

                <div className="space-y-2">
                  <label className="block text-sm font-medium text-slate-700">Pozadovane trilete obdobi cen</label>
                  <div className="w-full p-2.5 bg-slate-100 border border-slate-200 rounded-lg text-slate-700">{requiredPricePeriodLabel}</div>
                </div>

                <div className="space-y-2">
                  <label className="block text-sm font-medium text-slate-700">Nactene obdobi cen z CEN13A.xml</label>
                  <div className="w-full p-2.5 bg-slate-100 border border-slate-200 rounded-lg text-slate-700">{loadedPricePeriodLabel}</div>
                </div>

                <div className="space-y-2">
                  <label className="block text-sm font-medium text-slate-700">Nactene obdobi velikosti domu</label>
                  <div className="w-full p-2.5 bg-slate-100 border border-slate-200 rounded-lg text-slate-700">{loadedHousePeriodLabel}</div>
                </div>

                <div className="space-y-2">
                  <label className="block text-sm font-medium text-slate-700">Nactene obdobi velikosti bytu</label>
                  <div className="w-full p-2.5 bg-slate-100 border border-slate-200 rounded-lg text-slate-700">{loadedApartmentPeriodLabel}</div>
                </div>

                <div className="space-y-2 md:col-span-2">
                  <label className="block text-sm font-medium text-slate-700">Okres</label>
                  <input list="district-options" value={districtLabel} onChange={(e) => setDistrictLabel(e.target.value)} className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg" placeholder="Vyber nebo napis okres" />
                  <datalist id="district-options">
                    {(datasets.prices?.districtOptions ?? []).map((option) => (
                      <option key={option} value={option} />
                    ))}
                  </datalist>
                </div>

                <div className="space-y-2">
                  <label className="block text-sm font-medium text-slate-700">Krajovy fallback pro dum</label>
                  <input value={regionLabel} onChange={(e) => setRegionLabel(e.target.value)} className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg" placeholder="Napriklad Stredocesky" />
                </div>

                <div className="space-y-2">
                  <label className="block text-sm font-medium text-slate-700">Velikost obce</label>
                  <select value={municipalitySize} onChange={(e) => setMunicipalitySize(e.target.value)} className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg">
                    {MUNICIPALITY_OPTIONS.map((option) => (
                      <option key={option.value} value={option.value}>
                        {option.label}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
            </div>

            <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
              <h2 className="text-lg font-semibold border-b pb-4 mb-5 flex items-center gap-2">
                <Scale size={20} className="text-slate-500" />
                Statisticka hodnota a zakonne meze
              </h2>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
                <div className="space-y-2">
                  <label className="block text-sm font-medium text-slate-700">Prumerna velikost v okrese (m2)</label>
                  <div className="w-full p-2.5 bg-slate-100 border border-slate-200 rounded-lg text-slate-700">{formatNumber(autoDistrictArea, 2)}</div>
                  {(type === 'dum' || type === 'byt') && (
                    <div className="text-xs text-slate-500">
                      Krajovy fallback z {type === 'dum' ? '0140162502.xlsx' : '0140162507.xlsx'}: {formatNumber(regionalFallbackArea, 2)}
                    </div>
                  )}
                  <input type="number" step="0.01" value={districtAreaInput} onChange={(e) => setDistrictAreaInput(e.target.value)} className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg" placeholder="Rucni doplneni, pokud nemas okresni tabulku velikosti" />
                </div>

                <div className="space-y-2">
                  <label className="block text-sm font-medium text-slate-700">Kupni cena v okrese (Kc/m2)</label>
                  <div className="w-full p-2.5 bg-slate-100 border border-slate-200 rounded-lg text-slate-700">{formatCurrency(autoDistrictPrice)}</div>
                </div>

                <div className="space-y-2">
                  <label className="block text-sm font-medium text-slate-700">Republikova prumerna velikost (m2)</label>
                  <div className="w-full p-2.5 bg-slate-100 border border-slate-200 rounded-lg text-slate-700">{formatNumber(autoNationalArea, 2)}</div>
                  <input type="number" step="0.01" value={nationalAreaInput} onChange={(e) => setNationalAreaInput(e.target.value)} className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg" placeholder="Rucni doplneni, pokud chybi spravny narodní udaj" />
                </div>

                <div className="space-y-2">
                  <label className="block text-sm font-medium text-slate-700">Republikova kupni cena (Kc/m2)</label>
                  <div className="w-full p-2.5 bg-slate-100 border border-slate-200 rounded-lg text-slate-700">{formatCurrency(autoNationalPrice)}</div>
                </div>
              </div>
            </div>

            <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
              <h2 className="text-lg font-semibold border-b pb-4 mb-5 flex items-center gap-2">
                <TrendingUp size={20} className="text-slate-500" />
                Koeficient rustu cen
              </h2>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
                <div className="space-y-2">
                  <label className="block text-sm font-medium text-slate-700">Automaticky zvolene aktualni ctvrtleti</label>
                  <div className="w-full p-2.5 bg-slate-100 border border-slate-200 rounded-lg text-slate-700">{currentQuarterLabel}</div>
                </div>

                <div className="space-y-2">
                  <label className="block text-sm font-medium text-slate-700">Historicke srovnavaci ctvrtleti</label>
                  <div className="w-full p-2.5 bg-slate-100 border border-slate-200 rounded-lg text-slate-700">{historicQuarterLabel}</div>
                </div>

                <div className="space-y-2">
                  <label className="block text-sm font-medium text-slate-700">Aktualni index</label>
                  <div className="w-full p-2.5 bg-slate-100 border border-slate-200 rounded-lg text-slate-700">{formatNumber(currentIndex, 1)}</div>
                </div>

                <div className="space-y-2">
                  <label className="block text-sm font-medium text-slate-700">Historicky index</label>
                  <div className="w-full p-2.5 bg-slate-100 border border-slate-200 rounded-lg text-slate-700">{formatNumber(historicIndex, 1)}</div>
                </div>
              </div>
            </div>

            <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
              <h2 className="text-lg font-semibold border-b pb-4 mb-5 flex items-center gap-2">
                <Users size={20} className="text-slate-500" />
                Zohlednovane osoby
              </h2>
              <select value={persons} onChange={(e) => setPersons(e.target.value)} className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg">
                {PERSON_OPTIONS.map((option) => (
                  <option key={option.value} value={option.value}>
                    {option.label}
                  </option>
                ))}
              </select>
            </div>
          </div>

          <div className="lg:col-span-5 relative">
            <div className="sticky top-6 bg-blue-50/60 p-6 rounded-2xl shadow-sm border border-blue-200">
              <h2 className="text-xl font-bold text-blue-900 border-b border-blue-200 pb-4 mb-6">Vysledek vypoctu</h2>

              {calculation.missing.length > 0 && (
                <div className="mb-5 p-4 rounded-xl bg-amber-50 border border-amber-200 text-sm text-amber-900">
                  Pro plne metodicky vypocet jeste chybi: {calculation.missing.join(', ')}.
                </div>
              )}

              <div className="space-y-5">
                <div>
                  <div className="flex justify-between items-center mb-1">
                    <span className="text-sm font-semibold text-slate-600 uppercase tracking-wide">1. Hruba statisticka hodnota</span>
                    <span className="font-mono font-medium text-slate-800">{formatCurrency(calculation.raw)}</span>
                  </div>
                  <p className="text-xs text-slate-500">
                    {formatNumber(calculation.districtArea, 2)} m2 x {formatCurrency(autoDistrictPrice)}
                  </p>
                </div>

                <div>
                  <div className="flex justify-between items-center mb-1">
                    <span className="text-sm font-semibold text-slate-600 uppercase tracking-wide">2. Zakonny interval</span>
                    <span className="font-mono font-medium text-slate-800">
                      {formatCurrency(calculation.min)} az {formatCurrency(calculation.max)}
                    </span>
                  </div>
                </div>

                <div>
                  <div className="flex justify-between items-center mb-1">
                    <span className="text-sm font-semibold text-slate-600 uppercase tracking-wide">3. Statisticka hodnota po limitech</span>
                    <span className="font-mono font-medium text-slate-800">{formatCurrency(calculation.bounded)}</span>
                  </div>
                  <p className="text-xs text-slate-500">
                    {calculation.boundsApplied
                      ? 'Hruba hodnota byla upravena na zakonny limit.'
                      : 'Hruba hodnota zustala uvnitr zakonneho intervalu nebo limit nelze jeste spocitat.'}
                  </p>
                </div>

                <div>
                  <div className="flex justify-between items-center mb-1">
                    <span className="text-sm font-semibold text-slate-600 uppercase tracking-wide">4. Koeficient pozemku</span>
                    <span className="font-mono font-medium text-slate-800">x {formatNumber(calculation.kp, 2)}</span>
                  </div>
                </div>

                <div>
                  <div className="flex justify-between items-center mb-1">
                    <span className="text-sm font-semibold text-slate-600 uppercase tracking-wide">5. Koeficient rustu cen</span>
                    <span className="font-mono font-medium text-slate-800">x {formatNumber(calculation.growth, 2)}</span>
                  </div>
                  <p className="text-xs text-slate-500">{currentQuarterLabel} / {historicQuarterLabel}</p>
                </div>

                <div>
                  <div className="flex justify-between items-center mb-1">
                    <span className="text-sm font-semibold text-slate-600 uppercase tracking-wide">6. Koeficient osob</span>
                    <span className="font-mono font-medium text-slate-800">x {formatNumber(calculation.ko, 2)}</span>
                  </div>
                </div>

                <div className="pt-6 mt-6 border-t border-blue-200">
                  <p className="text-sm font-medium text-blue-800 mb-2">Hodnota chraneneho obydli</p>
                  <div className="text-4xl font-extrabold text-blue-900 tracking-tight">{formatCurrency(calculation.protectedValue)}</div>
                </div>
              </div>
            </div>
          </div>
        </div>

        <div className="bg-white p-6 md:p-8 rounded-2xl shadow-sm border border-slate-200">
          <div className="flex items-center gap-3 mb-6 border-b pb-4">
            <div className="p-2 bg-slate-100 text-slate-600 rounded-lg">
              <BookOpen size={24} />
            </div>
            <h2 className="text-xl font-bold text-slate-900">Metodicky stav</h2>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
            <section className="space-y-3">
              <h3 className="font-semibold text-slate-800">Co se doplnuje automaticky</h3>
              <ul className="text-sm text-slate-600 list-disc list-inside space-y-2">
                <li>okresni cena za m2 podle typu obydli a velikosti obce z CEN13A.xml,</li>
                <li>republikova cena za m2 z CEN13A.xml,</li>
                <li>ctvrtletni indexy podle data rozhodneho okamziku z 0140162508.xlsx a doplnkove z CEN15A1.xml a CEN15A1 (2023).xml,</li>
                <li>okresni, krajske a republikove velikosti rodinnych domu z 0140162502.xlsx.</li>
                <li>okresni, krajske a republikove velikosti bytu z 0140162507.xlsx.</li>
              </ul>
            </section>

            <section className="bg-amber-50 p-4 rounded-xl border border-amber-200">
              <h3 className="font-semibold text-amber-900 flex items-center gap-2 mb-2">
                <AlertCircle size={18} />
                Co jeste chybi
              </h3>
              {periodsAligned ? (
                <p className="text-sm text-amber-800 leading-relaxed">
                  Pri aktualnim datu jsou nactene soubory ve spravnem obdobi {requiredPricePeriodLabel}.
                </p>
              ) : (
                <p className="text-sm text-amber-800 leading-relaxed">
                  Pro uplnou metodickou automatiku musi vsechny nactene soubory odpovidat obdobi {requiredPricePeriodLabel}.
                  Pri nesouladu na to aplikace upozorni i ve vysledku vypoctu.
                </p>
              )}
            </section>
          </div>
        </div>
      </div>
    </div>
  );
}

export default App;
