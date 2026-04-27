#target photoshop
(function () {
	/**
	 * Runner für Photoshop (ExtendScript).
	 * Lädt die Arbeitsdatei "Export Pin Varianten.js" aus dem gleichen Ordner
	 * und führt sie aus.
	 */
	var runnerFile = new File($.fileName);
	var runnerFolder = runnerFile.parent;

	var workFile = new File(runnerFolder.fsName + "/export-pin-variants.js");
	if (!workFile.exists) {
		alert("Arbeitsdatei nicht gefunden:\n" + workFile.fsName);
		return;
	}

	$.evalFile(workFile);
})();
