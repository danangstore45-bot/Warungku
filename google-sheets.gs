const SHEET_NAME = "Transaksi";
const API_KEY = "GANTI_DENGAN_KUNCI_RAHASIA_UNIK";
const HEADERS = ["id", "type", "name", "category", "amount", "date", "time", "createdAt", "updatedAt", "deletedAt"];

function doGet(event) {
	const callback = event && event.parameter && event.parameter.callback;
	try {
		requireKey_(event.parameter.key);
		return respond_({ ok: true, records: readRecords_() }, callback);
	} catch (error) {
		return respond_({ ok: false, error: String(error.message || error), records: [] }, callback);
	}
}

function doPost(event) {
	try {
		const payload = JSON.parse(event.postData.contents || "{}");
		requireKey_(payload.key);
		if (payload.action !== "sync" || !Array.isArray(payload.records)) {
			throw new Error("Permintaan sinkronisasi tidak valid.");
		}
		writeRecords_(payload.records);
		return respond_({ ok: true });
	} catch (error) {
		return respond_({ ok: false, error: String(error.message || error) });
	}
}

function requireKey_(providedKey) {
	if (!API_KEY || API_KEY === "GANTI_DENGAN_KUNCI_RAHASIA_UNIK" || providedKey !== API_KEY) {
		throw new Error("Kunci akses tidak cocok. Periksa pengaturan Apps Script dan aplikasi.");
	}
}

function getSheet_() {
	const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
	if (!spreadsheet) throw new Error("Skrip harus dibuat dari Google Sheet tujuan.");
	let sheet = spreadsheet.getSheetByName(SHEET_NAME);
	if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
	if (sheet.getLastRow() === 0) {
		sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
	} else {
		const existingHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getDisplayValues()[0];
		if (HEADERS.some((header, index) => existingHeaders[index] !== header)) {
			throw new Error("Baris pertama sheet Transaksi harus berisi header bawaan skrip.");
		}
	}
	return sheet;
}

function readRecords_() {
	const sheet = getSheet_();
	if (sheet.getLastRow() < 2) return [];
	return sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues()
		.filter(row => row[0] !== "")
		.map(row => ({
			id: String(row[0]),
			type: String(row[1]),
			name: String(row[2]),
			category: String(row[3]),
			amount: Number(row[4]),
			date: row[5] instanceof Date ? Utilities.formatDate(row[5], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(row[5]),
			time: String(row[6]),
			createdAt: Number(row[7]) || 0,
			updatedAt: Number(row[8]) || 0,
			deletedAt: Number(row[9]) || 0
		}));
}

function writeRecords_(records) {
	const lock = LockService.getScriptLock();
	lock.waitLock(20000);
	try {
		const sheet = getSheet_();
		const existingRows = sheet.getLastRow() < 2
			? []
			: sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
		const rowById = new Map();
		existingRows.forEach((row, index) => {
			if (row[0] !== "") rowById.set(String(row[0]), { index: index + 2, row });
		});
		const rowsToAppend = [];

		records.forEach(record => {
			if (!record || typeof record.id !== "string" || !record.id) return;
			if (!["income", "expense"].includes(record.type) || !Number.isFinite(Number(record.amount))) return;
			const row = [
				record.id,
				record.type,
				String(record.name || ""),
				String(record.category || ""),
				Number(record.amount),
				String(record.date || ""),
				String(record.time || ""),
				Number(record.createdAt) || 0,
				Number(record.updatedAt) || 0,
				Number(record.deletedAt) || 0
			];
			const current = rowById.get(record.id);
			if (!current) {
				rowsToAppend.push(row);
				return;
			}
			const incomingVersion = Math.max(row[8], row[9]);
			const currentVersion = Math.max(Number(current.row[8]) || 0, Number(current.row[9]) || 0);
			const incomingIsDelete = row[9] > 0;
			const currentIsDelete = Number(current.row[9]) > 0;
			if (incomingVersion > currentVersion || (incomingVersion === currentVersion && incomingIsDelete && !currentIsDelete)) {
				sheet.getRange(current.index, 1, 1, HEADERS.length).setValues([row]);
			}
		});

		if (rowsToAppend.length) {
			sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, HEADERS.length).setValues(rowsToAppend);
		}
	} finally {
		lock.releaseLock();
	}
}

function respond_(data, callback) {
	if (callback && /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(callback)) {
		return ContentService.createTextOutput(callback + "(" + JSON.stringify(data) + ");")
			.setMimeType(ContentService.MimeType.JAVASCRIPT);
	}
	return ContentService.createTextOutput(JSON.stringify(data))
		.setMimeType(ContentService.MimeType.JSON);
}