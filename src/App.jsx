import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";

const SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets";
const STAT_FIELDS = [
  "seen_count",
  "correct_count",
  "wrong_count",
  "streak",
  "last_seen_at",
  "last_result",
  "mastery"
];
const REQUIRED_STATS_FIELDS = [...STAT_FIELDS];
const CARD_TEMPLATE_HEADERS = [
  "question",
  "answer",
  "pronunciation",
  "tags",
  "question_explanation",
  "answer_explanation"
];
const STATS_TEMPLATE_HEADERS = [
  "question",
  "answer",
  "pronunciation",
  "times_seen",
  "times_correct",
  "times_wrong",
  "streak",
  "last_seen_at",
  "last_result",
  "mastery"
];
const CARD_DATA_SHEET = "Card Data";
const CARD_STATS_SHEET = "Card Progress";
const CARD_COLUMN_ALIASES = {
  front: ["question"],
  back: ["answer"],
  romanization: ["pronunciation"],
  tags: ["tags"],
  question_explanation: ["question_explanation"],
  answer_explanation: ["answer_explanation"]
};
const STATS_COLUMN_ALIASES = {
  front: ["question"],
  back: ["answer"],
  romanization: ["pronunciation"],
  seen_count: ["times_seen"],
  correct_count: ["times_correct"],
  wrong_count: ["times_wrong"],
  streak: ["streak"],
  last_seen_at: ["last_seen_at"],
  last_result: ["last_result"],
  mastery: ["mastery"]
};
const STORAGE_KEYS = {
  sheetRef: "sheetCards.sheetRef",
  recentSheetRefs: "sheetCards.recentSheetRefs",
  recentSheetNames: "sheetCards.recentSheetNames"
};
const RECENT_SHEETS_LIMIT = 6;
const DEFAULT_CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID || "";

function getStoredValue(key, fallback = "") {
  try {
    const value = window.localStorage.getItem(key);
    if (value == null) {
      return fallback;
    }
    if (typeof value === "string" && value.trim() === "") {
      return fallback;
    }
    return value;
  } catch {
    return fallback;
  }
}

function getStoredList(key) {
  try {
    const raw = window.localStorage.getItem(key);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed.filter((value) => typeof value === "string" && value.trim());
  } catch {
    return [];
  }
}

function getStoredMap(key) {
  try {
    const raw = window.localStorage.getItem(key);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
      return {};
    }
    return Object.entries(parsed).reduce((acc, [mapKey, value]) => {
      if (typeof mapKey === "string" && typeof value === "string" && value.trim()) {
        acc[mapKey] = value.trim();
      }
      return acc;
    }, {});
  } catch {
    return {};
  }
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function parseNumber(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function normalizeKeyPart(value) {
  return String(value ?? "").trim().toLowerCase();
}

function buildQaBaseKey(front, back, romanization = "") {
  return [
    normalizeKeyPart(front),
    normalizeKeyPart(back),
    normalizeKeyPart(romanization)
  ].join("||");
}

function nextOccurrenceKey(baseKey, counterMap) {
  const next = (counterMap.get(baseKey) ?? 0) + 1;
  counterMap.set(baseKey, next);
  return `${baseKey}##${next}`;
}

function createDefaultStats() {
  return {
    seenCount: 0,
    correctCount: 0,
    wrongCount: 0,
    streak: 0,
    lastSeenAt: "",
    lastResult: "",
    mastery: 0
  };
}

function computeMastery(stats) {
  const seen = Math.max(0, stats.seenCount);
  const correct = Math.max(0, stats.correctCount);
  const wrong = Math.max(0, stats.wrongCount);
  const streak = Math.max(0, stats.streak);

  if (seen === 0) {
    return 0;
  }

  const accuracy = correct / seen;
  const streakBonus = Math.min(streak, 10) / 10;
  const wrongPenalty = wrong / seen;
  const score = accuracy * 0.72 + streakBonus * 0.28 - wrongPenalty * 0.2;

  return Number(clamp(score, 0, 1).toFixed(4));
}

function masteryToChoiceCount(mastery) {
  if (mastery < 0.4) return 2;
  if (mastery < 0.8) return 4;
  return 6;
}

function shuffle(list) {
  const arr = [...list];
  for (let i = arr.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

function parseTags(value) {
  if (!value) return [];
  return String(value)
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

function parseSpreadsheetId(input) {
  const value = input.trim();
  if (!value) return "";

  const match = value.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match) return match[1];

  if (/^[a-zA-Z0-9-_]{20,}$/.test(value)) {
    return value;
  }
  return "";
}

function colIndexToLetter(index) {
  let n = index + 1;
  let result = "";
  while (n > 0) {
    const remainder = (n - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

function quoteSheetTitle(sheetTitle) {
  const title = String(sheetTitle || CARD_DATA_SHEET).trim() || CARD_DATA_SHEET;
  return `'${title.replace(/'/g, "''")}'`;
}

function makeRange(sheetTitle, a1Range) {
  return `${quoteSheetTitle(sheetTitle)}!${a1Range}`;
}

function safeCell(row, index) {
  if (index == null || index < 0) return "";
  return row[index] ?? "";
}

function buildHeaderMap(headerRow) {
  const map = {};
  headerRow.forEach((rawName, index) => {
    const name = String(rawName ?? "").trim().toLowerCase();
    if (name) {
      map[name] = index;
    }
  });
  return map;
}

function buildAliasedColumnMap(headerRow, aliasesByField) {
  const headerMap = buildHeaderMap(headerRow);
  const normalized = {};
  Object.entries(aliasesByField).forEach(([field, aliases]) => {
    const matched = aliases.find((alias) => Number.isInteger(headerMap[alias]));
    if (matched) {
      normalized[field] = headerMap[matched];
    }
  });
  return normalized;
}

function pickWeightedCard(cards, excludeCardId) {
  const pool =
    cards.length > 1
      ? cards.filter((card) => card.cardId !== excludeCardId)
      : cards;

  if (pool.length === 0) {
    return null;
  }

  const weights = pool.map((card) => Math.max(0.05, 1.15 - card.mastery));
  const total = weights.reduce((sum, value) => sum + value, 0);
  let roll = Math.random() * total;

  for (let i = 0; i < pool.length; i += 1) {
    roll -= weights[i];
    if (roll <= 0) {
      return pool[i];
    }
  }
  return pool[pool.length - 1];
}

function pickDistractors(target, cards, count, getAnswer) {
  if (count <= 0) return [];

  const targetAnswer = String(getAnswer(target)).trim();

  const sameTag = cards.filter(
    (card) =>
      card.cardId !== target.cardId &&
      String(getAnswer(card)).trim() !== targetAnswer &&
      target.tags.some((tag) => card.tags.includes(tag))
  );

  const fallback = cards.filter(
    (card) =>
      card.cardId !== target.cardId &&
      String(getAnswer(card)).trim() !== targetAnswer
  );

  const selected = [];
  const usedAnswers = new Set([targetAnswer]);
  for (const card of shuffle(sameTag)) {
    if (selected.length >= count) break;
    const answer = String(getAnswer(card)).trim();
    if (usedAnswers.has(answer)) continue;
    selected.push(card);
    usedAnswers.add(answer);
  }

  for (const card of shuffle(fallback)) {
    if (selected.length >= count) break;
    const answer = String(getAnswer(card)).trim();
    if (usedAnswers.has(answer)) continue;
    selected.push(card);
    usedAnswers.add(answer);
  }

  return selected;
}

function formatPercent(value) {
  return `${Math.round(value * 100)}%`;
}

async function ensureGoogleIdentityScript() {
  if (window.google?.accounts?.oauth2) {
    return;
  }

  await new Promise((resolve, reject) => {
    const existing = document.querySelector('script[data-gis="true"]');
    if (existing) {
      existing.addEventListener("load", resolve, { once: true });
      existing.addEventListener(
        "error",
        () => reject(new Error("Failed to load Google Identity script.")),
        { once: true }
      );
      return;
    }

    const script = document.createElement("script");
    script.src = "https://accounts.google.com/gsi/client";
    script.async = true;
    script.defer = true;
    script.dataset.gis = "true";
    script.onload = resolve;
    script.onerror = () =>
      reject(new Error("Failed to load Google Identity script."));
    document.head.appendChild(script);
  });
}

async function sheetsGetValues({ spreadsheetId, range, accessToken }) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}`;
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`
    }
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Read failed (${response.status}): ${body}`);
  }

  return response.json();
}

async function sheetsBatchUpdate({ spreadsheetId, data, accessToken }) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values:batchUpdate`;
  const response = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      valueInputOption: "USER_ENTERED",
      data
    })
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Write failed (${response.status}): ${body}`);
  }
}

async function sheetsGetSpreadsheet({ spreadsheetId, accessToken }) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=properties.title,sheets.properties(sheetId,title)`;
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`
    }
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Metadata read failed (${response.status}): ${body}`);
  }

  return response.json();
}

async function sheetsSpreadsheetBatchUpdate({ spreadsheetId, requests, accessToken }) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`;
  const response = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ requests })
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Spreadsheet update failed (${response.status}): ${body}`);
  }
}

async function sheetsCreateSpreadsheet({ title, accessToken }) {
  const url = "https://sheets.googleapis.com/v4/spreadsheets";
  const response = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      properties: {
        title
      },
      sheets: [
        { properties: { title: CARD_DATA_SHEET } },
        { properties: { title: CARD_STATS_SHEET } }
      ]
    })
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Create sheet failed (${response.status}): ${body}`);
  }

  return response.json();
}

function buildStatUpdates(sheetTab, colByName, rowNumber, values) {
  const sheetPrefix = quoteSheetTitle(sheetTab);

  const fields = [];
  if (values.front != null && Number.isInteger(colByName.front)) {
    fields.push({ col: colByName.front, value: values.front });
  }
  if (values.back != null && Number.isInteger(colByName.back)) {
    fields.push({ col: colByName.back, value: values.back });
  }
  if (values.romanization != null && Number.isInteger(colByName.romanization)) {
    fields.push({ col: colByName.romanization, value: values.romanization });
  }

  const statValueByField = {
    seen_count: values.seenCount,
    correct_count: values.correctCount,
    wrong_count: values.wrongCount,
    streak: values.streak,
    last_seen_at: values.lastSeenAt,
    last_result: values.lastResult,
    mastery: Number(values.mastery.toFixed(4))
  };

  for (const field of STAT_FIELDS) {
    const col = colByName[field];
    if (!Number.isInteger(col)) continue;
    fields.push({ col, value: statValueByField[field] });
  }

  return fields.map(({ col, value }) => ({
    range: `${sheetPrefix}!${colIndexToLetter(col)}${rowNumber}`,
    majorDimension: "ROWS",
    values: [[value]]
  }));
}

export default function App() {
  const [appStage, setAppStage] = useState("connect");
  const [sheetRef, setSheetRef] = useState(() =>
    getStoredValue(STORAGE_KEYS.sheetRef, "")
  );
  const [recentSheetRefs, setRecentSheetRefs] = useState(() =>
    getStoredList(STORAGE_KEYS.recentSheetRefs)
  );
  const [recentSheetNames, setRecentSheetNames] = useState(() =>
    getStoredMap(STORAGE_KEYS.recentSheetNames)
  );
  const [newSheetTitle, setNewSheetTitle] = useState(() => {
    const date = new Date().toISOString().slice(0, 10);
    return `Sheet Cards ${date}`;
  });
  const [promptStudyNotes, setPromptStudyNotes] = useState("");
  const [studyMode, setStudyMode] = useState("front_only");
  const [showPronunciation, setShowPronunciation] = useState(true);
  const [currentDirection, setCurrentDirection] = useState("front_to_back");
  const [status, setStatus] = useState(
    "Connect Google to continue."
  );
  const [gisReady, setGisReady] = useState(false);
  const [accessToken, setAccessToken] = useState("");
  const [cards, setCards] = useState([]);
  const [currentCardId, setCurrentCardId] = useState("");
  const [choices, setChoices] = useState([]);
  const [answerState, setAnswerState] = useState(null);
  const [autoAdvanceMode, setAutoAdvanceMode] = useState("delay");
  const [autoAdvanceMs, setAutoAdvanceMs] = useState(1500);
  const [awaitingManualNext, setAwaitingManualNext] = useState(false);
  const [spokenTarget, setSpokenTarget] = useState("");
  const [pendingWrites, setPendingWrites] = useState(0);
  const [sessionAnswers, setSessionAnswers] = useState(0);
  const [sessionCorrect, setSessionCorrect] = useState(0);
  const [sessionWrongCards, setSessionWrongCards] = useState(0);
  const [sessionWrongSelections, setSessionWrongSelections] = useState(0);
  const [roundCompletedCount, setRoundCompletedCount] = useState(0);
  const [isFlushing, setIsFlushing] = useState(false);

  const tokenClientRef = useRef(null);
  const flushInFlightRef = useRef(false);
  const nextAdvanceTimerRef = useRef(null);
  const pendingNextRef = useRef(null);
  const speechSessionRef = useRef(0);
  const pendingStatsRef = useRef(new Map());
  const roundCompletedRef = useRef(new Set());
  const contextRef = useRef({
    spreadsheetId: "",
    statsSheet: {
      title: CARD_STATS_SHEET,
      colByName: {}
    }
  });

  const currentCard = useMemo(
    () => cards.find((card) => card.cardId === currentCardId) ?? null,
    [cards, currentCardId]
  );

  const sessionAccuracy = useMemo(() => {
    if (sessionAnswers === 0) return "0%";
    return formatPercent(sessionCorrect / sessionAnswers);
  }, [sessionAnswers, sessionCorrect]);
  const roundProgress = useMemo(() => {
    if (cards.length === 0) return "0 / 0";
    return `${roundCompletedCount} / ${cards.length}`;
  }, [cards.length, roundCompletedCount]);
  const tapAccuracy = useMemo(() => {
    const totalSelections = sessionAnswers + sessionWrongSelections;
    if (totalSelections === 0) return "0%";
    return formatPercent(sessionAnswers / totalSelections);
  }, [sessionAnswers, sessionWrongSelections]);
  const cardHintText = useMemo(() => {
    if (!currentCard) return "";
    if (showPronunciation && currentCard.romanization) {
      return currentCard.romanization;
    }
    return currentCard.tags.join(", ");
  }, [currentCard, showPronunciation]);

  const mostMissed = useMemo(() => {
    if (cards.length === 0) return "-";
    const missed = [...cards].sort((a, b) => b.wrongCount - a.wrongCount)[0];
    if (!missed || missed.wrongCount === 0) return "-";
    return `${missed.front} (${missed.wrongCount})`;
  }, [cards]);

  const spreadsheetId = useMemo(() => parseSpreadsheetId(sheetRef), [sheetRef]);
  const spreadsheetEditUrl = useMemo(() => {
    const raw = String(sheetRef || "").trim();
    if (/^https:\/\/docs\.google\.com\/spreadsheets\/d\//i.test(raw)) {
      return raw;
    }
    if (!spreadsheetId) return "";
    return `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;
  }, [sheetRef, spreadsheetId]);
  const csvPromptText = useMemo(() => {
    const headerRow = CARD_TEMPLATE_HEADERS.join("\t");
    const userMaterial = promptStudyNotes.trim() || "[Paste study material here]";
    return [
      "With the study material below, create tab-separated values (TSV) for my flashcard app.",
      `Use exactly this header row with TAB separators: ${headerRow}`,
      "Rules:",
      "- Return exactly one fenced code block only (no extra explanation text before or after).",
      "- Include the header row as the first line.",
      "- Use real TAB characters between columns and a new line for each row.",
      "- Keep this exact column order: question,answer,pronunciation,tags,question_explanation,answer_explanation.",
      "- pronunciation, tags, question_explanation, and answer_explanation are optional and may be blank.",
      "- Do not use tab characters inside cell values; replace inner tabs with a single space.",
      "- Keep each card on one line (no multi-line cells).",
      "- Do not output markdown tables.",
      "",
      "Study material:",
      userMaterial
    ].join("\n");
  }, [promptStudyNotes]);
  const clientId = DEFAULT_CLIENT_ID.trim();
  const isConfigured = Boolean(clientId);
  const promptFor = useCallback(
    (card, direction) => (direction === "front_to_back" ? card.front : card.back),
    []
  );
  const answerFor = useCallback(
    (card, direction) => (direction === "front_to_back" ? card.back : card.front),
    []
  );
  const promptExplanationFor = useCallback(
    (card, direction) =>
      direction === "front_to_back"
        ? card.questionExplanation
        : card.answerExplanation,
    []
  );
  const answerExplanationFor = useCallback(
    (card, direction) =>
      direction === "front_to_back"
        ? card.answerExplanation
        : card.questionExplanation,
    []
  );
  const resolveDirection = useCallback(() => {
    if (studyMode === "front_only") return "front_to_back";
    if (studyMode === "back_only") return "back_to_front";
    return Math.random() < 0.5 ? "front_to_back" : "back_to_front";
  }, [studyMode]);
  const speechSupported = useMemo(
    () =>
      typeof window !== "undefined" &&
      "speechSynthesis" in window &&
      "SpeechSynthesisUtterance" in window,
    []
  );

  const stopNarration = useCallback(() => {
    if (!speechSupported) return;
    speechSessionRef.current += 1;
    window.speechSynthesis.cancel();
    setSpokenTarget("");
  }, [speechSupported]);

  const speakText = useCallback(
    (text) => {
      const value = String(text ?? "").trim();
      if (!value) return;
      if (!speechSupported) {
        setStatus("Narration is not supported in this browser.");
        return;
      }
      window.speechSynthesis.cancel();
      const utterance = new window.SpeechSynthesisUtterance(value);
      utterance.lang = /[가-힣]/.test(value) ? "ko-KR" : "en-US";
      utterance.rate = 0.95;
      window.speechSynthesis.speak(utterance);
    },
    [speechSupported]
  );

  const speakSequence = useCallback(
    (items) => {
      if (!speechSupported) {
        setStatus("Narration is not supported in this browser.");
        return;
      }

      const entries = (items ?? [])
        .map((item) => ({
          target: item.target,
          text: String(item.text ?? "").trim()
        }))
        .filter((item) => item.text.length > 0);

      if (entries.length === 0) return;

      stopNarration();
      const sessionId = speechSessionRef.current + 1;
      speechSessionRef.current = sessionId;

      const speakAt = (index) => {
        if (speechSessionRef.current !== sessionId) return;
        if (index >= entries.length) {
          setSpokenTarget("");
          return;
        }

        const entry = entries[index];
        setSpokenTarget(entry.target);
        const utterance = new window.SpeechSynthesisUtterance(entry.text);
        utterance.lang = /[\uAC00-\uD7A3]/.test(entry.text) ? "ko-KR" : "en-US";
        utterance.rate = 0.95;
        utterance.onend = () => speakAt(index + 1);
        utterance.onerror = () => {
          setSpokenTarget("");
        };
        window.speechSynthesis.speak(utterance);
      };

      speakAt(0);
    },
    [speechSupported, stopNarration]
  );

  const clearPendingAdvance = useCallback(() => {
    if (nextAdvanceTimerRef.current) {
      window.clearTimeout(nextAdvanceTimerRef.current);
      nextAdvanceTimerRef.current = null;
    }
    pendingNextRef.current = null;
    setAwaitingManualNext(false);
  }, []);

  const refreshPendingCount = useCallback(() => {
    setPendingWrites(pendingStatsRef.current.size);
  }, []);

  const queueStatUpdate = useCallback(
    (card) => {
      pendingStatsRef.current.set(card.statsRowNumber, {
        seenCount: card.seenCount,
        correctCount: card.correctCount,
        wrongCount: card.wrongCount,
        streak: card.streak,
        lastSeenAt: card.lastSeenAt,
        lastResult: card.lastResult,
        mastery: card.mastery
      });
      refreshPendingCount();
    },
    [refreshPendingCount]
  );

  const rememberSheetRef = useCallback((value, sheetName = "") => {
    const nextId = parseSpreadsheetId(value);
    if (!nextId) return;
    setRecentSheetRefs((previous) => {
      const normalized = previous
        .map((item) => parseSpreadsheetId(item))
        .filter(Boolean);
      const updated = [nextId, ...normalized.filter((item) => item !== nextId)].slice(
        0,
        RECENT_SHEETS_LIMIT
      );
      return updated;
    });
    if (sheetName.trim()) {
      setRecentSheetNames((previous) => ({
        ...previous,
        [nextId]: sheetName.trim()
      }));
    }
  }, []);

  const resetRoundState = useCallback(() => {
    clearPendingAdvance();
    roundCompletedRef.current = new Set();
    setRoundCompletedCount(0);
    setSessionAnswers(0);
    setSessionCorrect(0);
    setSessionWrongCards(0);
    setSessionWrongSelections(0);
  }, [clearPendingAdvance]);

  const pickNextQuestion = useCallback((availableCards, previousCardId = "") => {
    clearPendingAdvance();
    stopNarration();
    if (availableCards.length === 0) {
      setCurrentCardId("");
      setChoices([]);
      return;
    }

    const remainingCards = availableCards.filter(
      (card) => !roundCompletedRef.current.has(card.cardId)
    );
    const candidateCards = remainingCards.length > 0 ? remainingCards : availableCards;
    const nextCard = pickWeightedCard(candidateCards, previousCardId);
    if (!nextCard) {
      setCurrentCardId("");
      setChoices([]);
      return;
    }

    const optionCount = Math.min(
      availableCards.length,
      masteryToChoiceCount(nextCard.mastery)
    );
    const nextDirection = resolveDirection();
    const distractors = pickDistractors(
      nextCard,
      availableCards,
      optionCount - 1,
      (card) => answerFor(card, nextDirection)
    );
    const nextChoices = shuffle([
      answerFor(nextCard, nextDirection),
      ...distractors.map((c) => answerFor(c, nextDirection))
    ]);

    setCurrentDirection(nextDirection);
    setCurrentCardId(nextCard.cardId);
    setChoices(nextChoices);
    setAnswerState(null);
  }, [answerFor, clearPendingAdvance, resolveDirection, stopNarration]);

  const flushPending = useCallback(
    async (silent = false) => {
      if (flushInFlightRef.current) return;
      if (!accessToken) return;
      if (pendingStatsRef.current.size === 0) {
        return;
      }

      const { spreadsheetId, statsSheet } = contextRef.current;
      if (!spreadsheetId) return;

      const data = [];

      for (const [rowNumber, values] of pendingStatsRef.current.entries()) {
        data.push(...buildStatUpdates(statsSheet.title, statsSheet.colByName, rowNumber, values));
      }

      if (data.length === 0) {
        return;
      }

      try {
        flushInFlightRef.current = true;
        setIsFlushing(true);
        await sheetsBatchUpdate({ spreadsheetId, data, accessToken });
        pendingStatsRef.current.clear();
        refreshPendingCount();
        if (!silent) {
          setStatus(`Synced ${data.length} update range(s) to Google Sheets.`);
        }
      } catch (error) {
        setStatus(error.message);
      } finally {
        flushInFlightRef.current = false;
        setIsFlushing(false);
      }
    },
    [accessToken, refreshPendingCount]
  );

  useEffect(() => {
    ensureGoogleIdentityScript()
      .then(() => setGisReady(true))
      .catch((error) => setStatus(error.message));
  }, []);

  useEffect(() => {
    try {
      window.localStorage.setItem(STORAGE_KEYS.sheetRef, sheetRef);
    } catch {
      // no-op
    }
  }, [sheetRef]);

  useEffect(() => {
    try {
      window.localStorage.setItem(
        STORAGE_KEYS.recentSheetRefs,
        JSON.stringify(recentSheetRefs)
      );
    } catch {
      // no-op
    }
  }, [recentSheetRefs]);

  useEffect(() => {
    try {
      window.localStorage.setItem(
        STORAGE_KEYS.recentSheetNames,
        JSON.stringify(recentSheetNames)
      );
    } catch {
      // no-op
    }
  }, [recentSheetNames]);

  useEffect(() => {
    const onVisibilityChange = () => {
      if (document.visibilityState === "hidden") {
        flushPending(true);
      }
    };
    document.addEventListener("visibilitychange", onVisibilityChange);
    return () => document.removeEventListener("visibilitychange", onVisibilityChange);
  }, [flushPending]);

  useEffect(() => {
    if (!accessToken) {
      setAppStage("connect");
      return;
    }
    setAppStage((previous) => (previous === "connect" ? "sheet" : previous));
  }, [accessToken]);

  useEffect(() => {
    if (!accessToken) return;
    const missingIds = recentSheetRefs.filter((id) => !recentSheetNames[id]);
    if (missingIds.length === 0) return;

    let cancelled = false;
    (async () => {
      const nameEntries = await Promise.all(
        missingIds.map(async (id) => {
          try {
            const metadata = await sheetsGetSpreadsheet({
              spreadsheetId: id,
              accessToken
            });
            const title = String(metadata.properties?.title || "").trim();
            if (!title) return null;
            return [id, title];
          } catch {
            return null;
          }
        })
      );

      if (cancelled) return;
      const updates = Object.fromEntries(nameEntries.filter(Boolean));
      if (Object.keys(updates).length === 0) return;
      setRecentSheetNames((previous) => ({
        ...previous,
        ...updates
      }));
    })();

    return () => {
      cancelled = true;
    };
  }, [accessToken, recentSheetNames, recentSheetRefs]);

  useEffect(() => {
    if (!accessToken) return;
    if (cards.length > 0) return;
    if (appStage === "study" || appStage === "summary") {
      setAppStage("sheet");
    }
  }, [accessToken, appStage, cards.length]);

  useEffect(() => {
    if (cards.length === 0) return;
    pickNextQuestion(cards);
  }, [studyMode]); // mode-only refresh

  useEffect(
    () => () => {
      clearPendingAdvance();
      stopNarration();
    },
    [clearPendingAdvance, stopNarration]
  );

  const handleSignIn = useCallback(() => {
    if (!gisReady) {
      setStatus(
        "Google login library is still loading or blocked. Disable script blockers and refresh."
      );
      return;
    }
    if (!clientId.trim()) {
      setStatus(
        "App setup incomplete: missing VITE_GOOGLE_CLIENT_ID. Ask the app owner to configure it and restart the app."
      );
      return;
    }

    try {
      tokenClientRef.current = window.google.accounts.oauth2.initTokenClient({
        client_id: clientId.trim(),
        scope: SHEETS_SCOPE,
        callback: (tokenResponse) => {
          if (tokenResponse.error) {
            setStatus(`Auth error: ${tokenResponse.error}`);
            return;
          }
          setAccessToken(tokenResponse.access_token);
          setAppStage("sheet");
          setStatus("Connected to Google. Next: load or create a sheet.");
        }
      });

      tokenClientRef.current.requestAccessToken({ prompt: "consent" });
    } catch (error) {
      setStatus(`Google sign-in failed to initialize: ${error.message}`);
    }
  }, [clientId, gisReady]);

  const handleInitializeSheetTemplate = useCallback(async () => {
    if (!accessToken) {
      setStatus("Connect Google first.");
      return;
    }
    if (!spreadsheetId) {
      setStatus("Invalid sheet URL or spreadsheet ID.");
      return;
    }

    const shouldContinue = window.confirm(
      `Initialize template in ${CARD_DATA_SHEET} and ${CARD_STATS_SHEET}? This rewrites row 1 headers in those tabs.`
    );
    if (!shouldContinue) {
      return;
    }

    try {
      setStatus(`Initializing ${CARD_DATA_SHEET} / ${CARD_STATS_SHEET} template...`);
      const metadata = await sheetsGetSpreadsheet({ spreadsheetId, accessToken });
      const sheetTitles = new Set(
        (metadata.sheets ?? [])
          .map((sheet) => sheet.properties?.title)
          .filter(Boolean)
      );

      const requests = [];
      if (!sheetTitles.has(CARD_DATA_SHEET)) {
        requests.push({
          addSheet: {
            properties: {
              title: CARD_DATA_SHEET
            }
          }
        });
      }
      if (!sheetTitles.has(CARD_STATS_SHEET)) {
        requests.push({
          addSheet: {
            properties: {
              title: CARD_STATS_SHEET
            }
          }
        });
      }

      if (requests.length > 0) {
        await sheetsSpreadsheetBatchUpdate({
          spreadsheetId,
          requests,
          accessToken
        });
      }

      const headerData = [
        {
          range: makeRange(
            CARD_DATA_SHEET,
            `A1:${colIndexToLetter(CARD_TEMPLATE_HEADERS.length - 1)}1`
          ),
          majorDimension: "ROWS",
          values: [CARD_TEMPLATE_HEADERS]
        },
        {
          range: makeRange(
            CARD_STATS_SHEET,
            `A1:${colIndexToLetter(STATS_TEMPLATE_HEADERS.length - 1)}1`
          ),
          majorDimension: "ROWS",
          values: [STATS_TEMPLATE_HEADERS]
        }
      ];
      await sheetsBatchUpdate({ spreadsheetId, data: headerData, accessToken });
      setStatus(
        `Template ready. Put card content rows in ${CARD_DATA_SHEET}, then click Load Cards.`
      );
    } catch (error) {
      setStatus(error.message);
    }
  }, [accessToken, spreadsheetId]);

  const handleCreateSheet = useCallback(async () => {
    if (!accessToken) {
      setStatus("Connect Google first.");
      return;
    }

    const title = String(newSheetTitle || "").trim() || `Sheet Cards ${new Date().toISOString().slice(0, 10)}`;

    try {
      setStatus("Creating a new spreadsheet...");
      const created = await sheetsCreateSpreadsheet({ title, accessToken });
      const createdId = created.spreadsheetId;
      const createdUrl =
        created.spreadsheetUrl || `https://docs.google.com/spreadsheets/d/${createdId}/edit`;

      const headerData = [
        {
          range: makeRange(
            CARD_DATA_SHEET,
            `A1:${colIndexToLetter(CARD_TEMPLATE_HEADERS.length - 1)}1`
          ),
          majorDimension: "ROWS",
          values: [CARD_TEMPLATE_HEADERS]
        },
        {
          range: makeRange(
            CARD_STATS_SHEET,
            `A1:${colIndexToLetter(STATS_TEMPLATE_HEADERS.length - 1)}1`
          ),
          majorDimension: "ROWS",
          values: [STATS_TEMPLATE_HEADERS]
        }
      ];
      await sheetsBatchUpdate({
        spreadsheetId: createdId,
        data: headerData,
        accessToken
      });

      setSheetRef(createdUrl);
      rememberSheetRef(createdUrl, title);
      setStatus(`Created new sheet template: ${title}. Add cards, then click Load Cards.`);
    } catch (error) {
      setStatus(error.message);
    }
  }, [accessToken, newSheetTitle, rememberSheetRef]);

  const handleLoadCards = useCallback(async (options = {}) => {
    const targetSheetRef = String(options.sheetRef ?? sheetRef);
    const targetSpreadsheetId =
      String(options.spreadsheetId ?? parseSpreadsheetId(targetSheetRef));

    if (!accessToken) {
      setStatus("Connect Google first.");
      return;
    }
    if (!targetSpreadsheetId) {
      setStatus("Invalid sheet URL or spreadsheet ID.");
      return;
    }

    try {
      setStatus(`Loading cards from ${CARD_DATA_SHEET} and stats from ${CARD_STATS_SHEET}...`);
      const [cardsResponse, statsResponse, metadata] = await Promise.all([
        sheetsGetValues({
          spreadsheetId: targetSpreadsheetId,
          range: makeRange(CARD_DATA_SHEET, "A:Z"),
          accessToken
        }),
        sheetsGetValues({
          spreadsheetId: targetSpreadsheetId,
          range: makeRange(CARD_STATS_SHEET, "A:Z"),
          accessToken
        }),
        sheetsGetSpreadsheet({
          spreadsheetId: targetSpreadsheetId,
          accessToken
        })
      ]);

      const cardRows = cardsResponse.values ?? [];
      if (cardRows.length === 0) {
        setCards([]);
        setCurrentCardId("");
        setChoices([]);
        setStatus(
          `${CARD_DATA_SHEET} is empty. Initialize the template if needed, then add card rows under the header.`
        );
        return;
      }

      const cardsColByName = buildAliasedColumnMap(cardRows[0], CARD_COLUMN_ALIASES);
      const requiredCardColumns = ["front", "back"];
      const missingCardColumns = requiredCardColumns.filter(
        (name) => !Number.isInteger(cardsColByName[name])
      );
      if (missingCardColumns.length > 0) {
        setStatus(
          `${CARD_DATA_SHEET} is missing required columns: ${missingCardColumns.join(", ")}. Click "Initialize Sheet Template".`
        );
        return;
      }

      const statsRows = statsResponse.values ?? [];
      if (statsRows.length === 0) {
        setStatus(
          `${CARD_STATS_SHEET} is empty. Click "Initialize Sheet Template" first.`
        );
        return;
      }

      const statsColByName = buildAliasedColumnMap(statsRows[0], STATS_COLUMN_ALIASES);
      const missingStatsColumns = REQUIRED_STATS_FIELDS.filter(
        (name) => !Number.isInteger(statsColByName[name])
      );
      if (missingStatsColumns.length > 0) {
        setStatus(
          `${CARD_STATS_SHEET} is missing required columns: ${missingStatsColumns.join(", ")}. Click "Initialize Sheet Template".`
        );
        return;
      }
      const hasStatsQaColumns =
        Number.isInteger(statsColByName.front) &&
        Number.isInteger(statsColByName.back) &&
        Number.isInteger(statsColByName.romanization);
      if (!hasStatsQaColumns) {
        setStatus(
          `${CARD_STATS_SHEET} must include question, answer, and pronunciation columns. Click "Initialize Sheet Template".`
        );
        return;
      }

      pendingStatsRef.current.clear();

      const statsByKey = new Map();
      const qaStatsCounter = new Map();
      for (let rowIndex = 1; rowIndex < statsRows.length; rowIndex += 1) {
        const row = statsRows[rowIndex];
        const statsFront = String(safeCell(row, statsColByName.front)).trim();
        const statsBack = String(safeCell(row, statsColByName.back)).trim();
        const statsRomanization = String(safeCell(row, statsColByName.romanization)).trim();
        if (!statsFront && !statsBack) continue;

        const baseStats = createDefaultStats();
        baseStats.seenCount = parseNumber(safeCell(row, statsColByName.seen_count));
        baseStats.correctCount = parseNumber(safeCell(row, statsColByName.correct_count));
        baseStats.wrongCount = parseNumber(safeCell(row, statsColByName.wrong_count));
        baseStats.streak = parseNumber(safeCell(row, statsColByName.streak));
        baseStats.lastSeenAt = String(safeCell(row, statsColByName.last_seen_at)).trim();
        baseStats.lastResult = String(safeCell(row, statsColByName.last_result)).trim();
        baseStats.mastery = parseNumber(safeCell(row, statsColByName.mastery));
        if (baseStats.mastery <= 0 && baseStats.seenCount > 0) {
          baseStats.mastery = computeMastery(baseStats);
        }

        const matchKey = nextOccurrenceKey(
          `qa:${buildQaBaseKey(statsFront, statsBack, statsRomanization)}`,
          qaStatsCounter
        );

        statsByKey.set(matchKey, {
          rowNumber: rowIndex + 1,
          ...baseStats
        });
      }

      let nextStatsRowNumber = Math.max(2, statsRows.length + 1);
      const qaCardCounter = new Map();
      const nextCards = [];
      for (let rowIndex = 1; rowIndex < cardRows.length; rowIndex += 1) {
        const row = cardRows[rowIndex];
        const front = String(safeCell(row, cardsColByName.front)).trim();
        const back = String(safeCell(row, cardsColByName.back)).trim();
        const romanization = String(safeCell(row, cardsColByName.romanization)).trim();
        if (!front && !back) {
          continue;
        }

        const rowNumber = rowIndex + 1;
        const matchKey = nextOccurrenceKey(
          `qa:${buildQaBaseKey(front, back, romanization)}`,
          qaCardCounter
        );

        let stats = statsByKey.get(matchKey);
        if (!stats) {
          const baseStats = createDefaultStats();
          stats = {
            rowNumber: nextStatsRowNumber,
            ...baseStats
          };
          nextStatsRowNumber += 1;
          statsByKey.set(matchKey, stats);
          pendingStatsRef.current.set(stats.rowNumber, {
            ...baseStats,
            front,
            back,
            romanization
          });
        }

        nextCards.push({
          contentRowNumber: rowNumber,
          statsRowNumber: stats.rowNumber,
          cardId: matchKey,
          front,
          back,
          romanization,
          tags: parseTags(safeCell(row, cardsColByName.tags)),
          questionExplanation: String(
            safeCell(row, cardsColByName.question_explanation)
          ).trim(),
          answerExplanation: String(
            safeCell(row, cardsColByName.answer_explanation)
          ).trim(),
          seenCount: stats.seenCount,
          correctCount: stats.correctCount,
          wrongCount: stats.wrongCount,
          streak: stats.streak,
          lastSeenAt: stats.lastSeenAt,
          lastResult: stats.lastResult,
          mastery: stats.mastery
        });
      }

      contextRef.current = {
        spreadsheetId: targetSpreadsheetId,
        statsSheet: {
          title: CARD_STATS_SHEET,
          colByName: statsColByName
        }
      };

      setCards(nextCards);
      resetRoundState();
      pickNextQuestion(nextCards);
      refreshPendingCount();
      rememberSheetRef(targetSheetRef || targetSpreadsheetId, String(metadata.properties?.title || ""));
      setAppStage("sheet");

      if (pendingStatsRef.current.size > 0) {
        await flushPending(true);
      }

      setStatus(`Loaded ${nextCards.length} cards. Click Start Study Round.`);
    } catch (error) {
      const message = String(error.message ?? "");
      if (message.includes("Unable to parse range")) {
        setStatus(
          `Sheet tabs are missing. Click "Initialize Sheet Template" to create ${CARD_DATA_SHEET} and ${CARD_STATS_SHEET}.`
        );
        return;
      }
      setStatus(message);
    }
  }, [
    accessToken,
    flushPending,
    pickNextQuestion,
    rememberSheetRef,
    resetRoundState,
    refreshPendingCount,
    sheetRef,
    spreadsheetId
  ]);

  const handleSelectRecentSheet = useCallback(
    (id) => {
      const nextUrl = `https://docs.google.com/spreadsheets/d/${id}/edit`;
      setSheetRef(nextUrl);
      setStatus(`Selected ${recentSheetNames[id] || "sheet"}. Auto-loading cards...`);
      handleLoadCards({
        spreadsheetId: id,
        sheetRef: nextUrl
      });
    },
    [handleLoadCards, recentSheetNames]
  );

  const handleUnloadCards = useCallback(async () => {
    if (pendingWrites > 0) {
      const shouldSync = window.confirm(
        "You have pending updates. Sync before unloading cards?"
      );
      if (shouldSync) {
        await flushPending(false);
      }
    }

    clearPendingAdvance();
    stopNarration();
    setCards([]);
    setCurrentCardId("");
    setChoices([]);
    setAnswerState(null);
    resetRoundState();
    setAppStage("sheet");
    setStatus("Cards unloaded. Load a sheet to study again.");
  }, [clearPendingAdvance, flushPending, pendingWrites, resetRoundState, stopNarration]);

  const goToQueuedNextCard = useCallback(() => {
    const queued = pendingNextRef.current;
    if (!queued) return;
    clearPendingAdvance();
    pickNextQuestion(queued.cards, queued.previousCardId);
  }, [clearPendingAdvance, pickNextQuestion]);

  const handleStartStudyRound = useCallback(() => {
    if (cards.length === 0) {
      setStatus("Load cards first.");
      return;
    }
    resetRoundState();
    pickNextQuestion(cards);
    setAppStage("study");
    setStatus("Study round started.");
  }, [cards, pickNextQuestion, resetRoundState]);

  const handleOpenSheet = useCallback(() => {
    if (!spreadsheetEditUrl) {
      setStatus("Enter a valid sheet URL or spreadsheet ID first.");
      return;
    }

    const opened = window.open(spreadsheetEditUrl, "_blank", "noopener,noreferrer");
    if (!opened) {
      setStatus("Browser blocked the new tab. Allow popups for this site.");
    }
  }, [spreadsheetEditUrl]);

  const handleCopyCsvPrompt = useCallback(async () => {
    try {
      if (navigator.clipboard?.writeText) {
        await navigator.clipboard.writeText(csvPromptText);
      } else {
        const temp = document.createElement("textarea");
        temp.value = csvPromptText;
        temp.setAttribute("readonly", "true");
        temp.style.position = "fixed";
        temp.style.top = "-9999px";
        document.body.appendChild(temp);
        temp.select();
        document.execCommand("copy");
        document.body.removeChild(temp);
      }
      setStatus("Prompt copied. Paste it into ChatGPT and run it.");
    } catch {
      setStatus("Copy failed. Select the prompt text and copy manually.");
    }
  }, [csvPromptText]);

  const handleAnswer = useCallback(
    (choice) => {
      if (!currentCard) return;
      const correctAnswer = answerFor(currentCard, currentDirection);
      const applyCardCompletion = (cardsSnapshot, hadMistake) => {
        const now = new Date().toISOString();
        const updatedCard = {
          ...currentCard,
          seenCount: currentCard.seenCount + 1,
          correctCount: currentCard.correctCount + (hadMistake ? 0 : 1),
          wrongCount: currentCard.wrongCount + (hadMistake ? 1 : 0),
          streak: hadMistake ? 0 : currentCard.streak + 1,
          lastSeenAt: now,
          lastResult: hadMistake ? "wrong" : "correct"
        };
        updatedCard.mastery = computeMastery(updatedCard);

        const updatedCards = cardsSnapshot.map((card) =>
          card.cardId === updatedCard.cardId ? updatedCard : card
        );
        setCards(updatedCards);
        queueStatUpdate(updatedCard);

        const nextAnswers = sessionAnswers + 1;
        const nextCorrect = sessionCorrect + (hadMistake ? 0 : 1);
        setSessionAnswers(nextAnswers);
        setSessionCorrect(nextCorrect);
        if (hadMistake) {
          setSessionWrongCards((prev) => prev + 1);
        }
        setAnswerState({
          isCorrect: !hadMistake,
          completedWithMistake: hadMistake,
          expected: correctAnswer,
          requiresCorrection: false
        });

        if (!roundCompletedRef.current.has(updatedCard.cardId)) {
          roundCompletedRef.current.add(updatedCard.cardId);
          setRoundCompletedCount(roundCompletedRef.current.size);
        }

        if (nextAnswers % 10 === 0) {
          flushPending(true);
        }
        const completedRound = roundCompletedRef.current.size >= updatedCards.length;
        if (completedRound) {
          clearPendingAdvance();
          stopNarration();
          pendingNextRef.current = null;
          setAwaitingManualNext(false);
          setAppStage("summary");
          setStatus("Round complete. Review summary and start the next round.");
          return;
        }

        pendingNextRef.current = {
          cards: updatedCards,
          previousCardId: currentCard.cardId
        };

        if (autoAdvanceMode === "manual") {
          setAwaitingManualNext(true);
          return;
        }

        nextAdvanceTimerRef.current = window.setTimeout(() => {
          goToQueuedNextCard();
        }, autoAdvanceMs);
      };

      if (answerState?.requiresCorrection) {
        if (choice !== correctAnswer) {
          return;
        }
        applyCardCompletion(cards, true);
        return;
      }

      if (answerState) {
        return;
      }

      if (choice !== correctAnswer) {
        setSessionWrongSelections((prev) => prev + 1);
        setAnswerState({
          isCorrect: false,
          expected: correctAnswer,
          requiresCorrection: true,
          wrongChoice: choice,
          completedWithMistake: true
        });
        return;
      }

      applyCardCompletion(cards, false);
    },
    [
      answerState,
      answerFor,
      autoAdvanceMode,
      autoAdvanceMs,
      cards,
      clearPendingAdvance,
      currentDirection,
      currentCard,
      flushPending,
      goToQueuedNextCard,
      queueStatUpdate,
      sessionWrongCards,
      stopNarration,
      sessionAnswers,
      sessionCorrect
    ]
  );

  const isCorrectionPhase = Boolean(answerState?.requiresCorrection);
  const feedbackTone =
    answerState?.requiresCorrection || answerState?.completedWithMistake
      ? "bad"
      : "good";
  const speakQuestion = useCallback(() => {
    if (!currentCard) return;
    const prompt = promptFor(currentCard, currentDirection);
    speakSequence([{ target: "question", text: prompt }]);
  }, [currentCard, currentDirection, promptFor, speakSequence]);
  const speakAnswers = useCallback(() => {
    if (!currentCard) return;
    speakSequence(
      choices.map((choice, index) => ({
        target: `choice-${index}`,
        text: choice
      }))
    );
  }, [choices, currentCard, speakSequence]);

  const canGoSheet = Boolean(accessToken);
  const canGoStudy = cards.length > 0;

  return (
    <div className="page">
      <header className="hero">
        <h1>Sheet Cards</h1>
        <p>Study flow: Connect Google, then Load Sheet, then Study, then Round Summary.</p>
        <div className="hero-actions">
          <button
            className={`btn ${appStage === "connect" ? "btn-accent" : "btn-subtle"}`}
            onClick={() => setAppStage("connect")}
          >
            Home
          </button>
          {canGoSheet && (
            <button
              className={`btn ${appStage === "sheet" ? "btn-accent" : "btn-subtle"}`}
              onClick={() => setAppStage("sheet")}
            >
              Sheet
            </button>
          )}
          {canGoStudy && (
            <button
              className={`btn ${appStage === "study" ? "btn-accent" : "btn-subtle"}`}
              onClick={() => setAppStage("study")}
            >
              Study
            </button>
          )}
          {canGoStudy && (
            <button
              className={`btn ${appStage === "summary" ? "btn-accent" : "btn-subtle"}`}
              onClick={() => setAppStage("summary")}
            >
              Stats
            </button>
          )}
        </div>
      </header>

      {appStage === "connect" && (
        <section className="panel stage-panel">
          <h2>Connect Google</h2>
          <p className="status-inline">
            {isConfigured
              ? "Sign in once to allow read/write access to your sheets."
              : "Missing VITE_GOOGLE_CLIENT_ID in app config."}
          </p>
          <div className="actions">
            <button className="btn btn-accent" onClick={handleSignIn} disabled={!isConfigured}>
              {accessToken ? "Reconnect Google" : "Connect Google"}
            </button>
            <button className="btn btn-subtle" onClick={() => setAppStage("sheet")} disabled={!accessToken}>
              Continue to Sheet
            </button>
          </div>
        </section>
      )}

      {appStage === "sheet" && (
        <section className="panel stage-panel">
          <h2>Load Or Create Sheet</h2>
          <div className="field-grid compact">
            <label className="field">
              <span>Sheet URL or Spreadsheet ID</span>
              <input
                type="text"
                value={sheetRef}
                onChange={(event) => setSheetRef(event.target.value)}
                placeholder="https://docs.google.com/spreadsheets/d/..."
              />
            </label>
            <div className="field fixed-tabs">
              <span>Expected Tabs</span>
              <p>{`${CARD_DATA_SHEET} + ${CARD_STATS_SHEET}`}</p>
            </div>
          </div>
          {recentSheetRefs.length > 0 && (
            <div className="recent-row">
              <span>Recent Sheets</span>
              <div className="recent-list">
                {recentSheetRefs.map((id) => (
                  <button
                    key={id}
                    className={`btn recent-btn ${spreadsheetId === id ? "btn-accent" : "btn-subtle"}`}
                    onClick={() => handleSelectRecentSheet(id)}
                    title={id}
                  >
                    {recentSheetNames[id] || "Loading name..."}
                  </button>
                ))}
              </div>
            </div>
          )}
          <div className="actions">
            <button
              className="btn"
              onClick={handleInitializeSheetTemplate}
              disabled={!accessToken || !spreadsheetId}
            >
              Initialize Sheet Template
            </button>
            <button
              className="btn btn-accent"
              onClick={handleLoadCards}
              disabled={!accessToken || !spreadsheetId}
            >
              Load Cards
            </button>
            <button
              className="btn btn-subtle"
              onClick={handleUnloadCards}
              disabled={cards.length === 0}
            >
              Unload Cards
            </button>
            <button
              className="btn btn-subtle"
              onClick={() => flushPending(false)}
              disabled={pendingWrites === 0 || isFlushing}
            >
              {isFlushing ? "Syncing..." : "Sync Pending"}
            </button>
            <button
              className="btn btn-subtle"
              onClick={handleOpenSheet}
              disabled={!spreadsheetId}
            >
              Open Sheet
            </button>
          </div>
          <div className="new-sheet">
            <label className="field">
              <span>New Sheet Name</span>
              <input
                type="text"
                value={newSheetTitle}
                onChange={(event) => setNewSheetTitle(event.target.value)}
                placeholder="Sheet Cards"
              />
            </label>
            <div className="actions">
              <button className="btn" onClick={handleCreateSheet} disabled={!accessToken}>
                Create New Sheet
              </button>
              <button className="btn btn-accent" onClick={handleStartStudyRound} disabled={cards.length === 0}>
                Start Study Round
              </button>
            </div>
          </div>
          <div className="prompt-builder">
            <h3>Prompt Starter (TSV For Sheets)</h3>
            <p className="status-inline">
              Paste your topic/study material below, then copy this prompt into ChatGPT. The output will be tab-separated for direct paste into Google Sheets.
            </p>
            <label className="field">
              <span>Study Material Notes</span>
              <textarea
                value={promptStudyNotes}
                onChange={(event) => setPromptStudyNotes(event.target.value)}
                rows={5}
                placeholder="Example: Korean travel vocabulary for beginners..."
              />
            </label>
            <label className="field">
              <span>Copyable Prompt (TSV)</span>
              <textarea value={csvPromptText} readOnly rows={14} />
            </label>
            <div className="actions">
              <button className="btn btn-subtle" onClick={handleCopyCsvPrompt}>
                Copy Prompt
              </button>
            </div>
          </div>
          <section className="metrics-panel compact-metrics">
            <div className="metric">
              <span>Cards Loaded</span>
              <strong>{cards.length}</strong>
            </div>
            <div className="metric">
              <span>Pending Writes</span>
              <strong>{pendingWrites}</strong>
            </div>
          </section>
        </section>
      )}

      {appStage === "study" && (
        <section className="panel study-panel">
          {!currentCard ? (
            <div className="empty-state">Start a study round from the Sheet screen.</div>
          ) : (
            <div className="study-state">
              <div className="card-meta">
                <span>{`Round ${roundProgress}`}</span>
                <span>{`First Try ${sessionAccuracy}`}</span>
                <span>{`Missed ${sessionWrongCards}`}</span>
                <span>{`Mastery ${currentCard.mastery.toFixed(2)}`}</span>
              </div>
              <div className="actions card-actions">
                <button
                  className={`btn ${studyMode === "front_only" ? "btn-accent" : "btn-subtle"}`}
                  onClick={() => setStudyMode("front_only")}
                >
                  Front Only
                </button>
                <button
                  className={`btn ${studyMode === "back_only" ? "btn-accent" : "btn-subtle"}`}
                  onClick={() => setStudyMode("back_only")}
                >
                  Back Only
                </button>
                <button
                  className={`btn ${studyMode === "random" ? "btn-accent" : "btn-subtle"}`}
                  onClick={() => setStudyMode("random")}
                >
                  Random
                </button>
                <button
                  className={`btn ${showPronunciation ? "btn-accent" : "btn-subtle"}`}
                  onClick={() => setShowPronunciation((value) => !value)}
                >
                  {showPronunciation ? "Pronunciation On" : "Pronunciation Off"}
                </button>
              </div>
              <div className="actions card-actions">
                <button
                  className={`btn ${autoAdvanceMode === "delay" ? "btn-accent" : "btn-subtle"}`}
                  onClick={() => setAutoAdvanceMode("delay")}
                >
                  Auto Next
                </button>
                <button
                  className={`btn ${autoAdvanceMode === "manual" ? "btn-accent" : "btn-subtle"}`}
                  onClick={() => setAutoAdvanceMode("manual")}
                >
                  Manual Next
                </button>
                {autoAdvanceMode === "delay" && (
                  <label className="field inline-field">
                    <span>Show Answer</span>
                    <select
                      value={String(autoAdvanceMs)}
                      onChange={(event) => setAutoAdvanceMs(Number(event.target.value))}
                    >
                      <option value="900">0.9s</option>
                      <option value="1500">1.5s</option>
                      <option value="2500">2.5s</option>
                      <option value="4000">4.0s</option>
                    </select>
                  </label>
                )}
              </div>
              <div className="actions card-actions">
                <button className="btn btn-subtle" onClick={speakQuestion} disabled={!speechSupported}>
                  Read Question
                </button>
                <button
                  className="btn btn-subtle"
                  onClick={speakAnswers}
                  disabled={!speechSupported}
                >
                  Read Answers
                </button>
                <button className="btn btn-subtle" onClick={stopNarration} disabled={!speechSupported}>
                  Stop Voice
                </button>
                <button className="btn btn-subtle" onClick={() => setAppStage("summary")}>
                  End Round
                </button>
              </div>
              <h2
                className={[
                  "card-front",
                  spokenTarget === "question" ? "reading-focus" : ""
                ]
                  .filter(Boolean)
                  .join(" ")}
              >
                {promptFor(currentCard, currentDirection)}
              </h2>
              <p className="card-hint">
                {cardHintText}
              </p>
              <div className="choices">
                {choices.map((choice, index) => {
                  const correctChoice = answerFor(currentCard, currentDirection);
                  const isSelectedWrong =
                    isCorrectionPhase && choice === answerState?.wrongChoice;
                  const isAnswer = answerState && choice === correctChoice;

                  return (
                    <button
                      key={`${currentCard.cardId}_${choice}`}
                      className={[
                        "choice",
                        isAnswer ? "choice-correct" : "",
                        isSelectedWrong ? "choice-muted" : "",
                        spokenTarget === `choice-${index}` ? "reading-focus" : ""
                      ]
                        .filter(Boolean)
                        .join(" ")}
                      onClick={() => handleAnswer(choice)}
                      disabled={isCorrectionPhase ? choice !== correctChoice : Boolean(answerState)}
                    >
                      {choice}
                    </button>
                  );
                })}
              </div>
              {answerState && (
                <p className={`feedback ${feedbackTone}`}>
                  {answerState.requiresCorrection
                    ? "Not yet. Tap the correct answer to continue."
                    : answerState.completedWithMistake
                      ? "Completed after correction. Counted as wrong."
                      : "Correct on first try."}
                </p>
              )}
              {answerState && promptExplanationFor(currentCard, currentDirection) && (
                <p className="card-explain">
                  <strong>Question Explanation:</strong>{" "}
                  {promptExplanationFor(currentCard, currentDirection)}
                </p>
              )}
              {answerState && answerExplanationFor(currentCard, currentDirection) && (
                <p className="card-explain">
                  <strong>Answer Explanation:</strong>{" "}
                  {answerExplanationFor(currentCard, currentDirection)}
                </p>
              )}
              {awaitingManualNext && (
                <div className="actions card-actions">
                  <button className="btn btn-accent" onClick={goToQueuedNextCard}>
                    Next Card
                  </button>
                </div>
              )}
            </div>
          )}
        </section>
      )}

      {appStage === "summary" && (
        <section className="panel stage-panel">
          <h2>Stats</h2>
          <section className="metrics-panel">
            <div className="metric">
              <span>Cards Completed</span>
              <strong>{roundProgress}</strong>
            </div>
            <div className="metric">
              <span>First Try Accuracy</span>
              <strong>{sessionAccuracy}</strong>
            </div>
            <div className="metric">
              <span>Tap Accuracy</span>
              <strong>{tapAccuracy}</strong>
            </div>
            <div className="metric">
              <span>Cards Missed</span>
              <strong>{sessionWrongCards}</strong>
            </div>
            <div className="metric">
              <span>Wrong Selections</span>
              <strong>{sessionWrongSelections}</strong>
            </div>
            <div className="metric">
              <span>Most Missed</span>
              <strong>{mostMissed}</strong>
            </div>
            <div className="metric">
              <span>Pending Writes</span>
              <strong>{pendingWrites}</strong>
            </div>
          </section>
          <div className="actions">
            <button
              className="btn btn-subtle"
              onClick={() => flushPending(false)}
              disabled={pendingWrites === 0 || isFlushing}
            >
              {isFlushing ? "Syncing..." : "Sync Pending"}
            </button>
            <button className="btn" onClick={() => setAppStage("sheet")}>
              Back To Sheet
            </button>
            <button className="btn btn-accent" onClick={handleStartStudyRound} disabled={cards.length === 0}>
              Start Next Round
            </button>
          </div>
        </section>
      )}

      <p className="status">{status}</p>
    </div>
  );
}
