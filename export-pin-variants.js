
(function () {
	/**
	 * Exportiert pro Bilder-Untergruppe alle Varianten aus Varianten-Untergruppen als JPG nach ~/Downloads.
	 *
	 * Erwartete Struktur (Top-Level):
	 * - "Bilder" (Gruppe)
	 * 		- "Bild 1" (Untergruppe)
	 * 		- "Bild 2" (Untergruppe)
	 * - "Varianten" (Gruppe)
	 * 		- "t1_v6-10" (Untergruppe; Muster: t<zahl>_v<zahl>-<zahl>)
	 * 			- Ebene "Overlay-Hintergrund" (Formebene, Solid Fill)
	 * 			- 1..n Text-Layer (beliebig verschachtelt)
	 *
	 * Dateiname:
	 * <docPrefix> + <variantPrefix> + <laufnummer> + "_p" + <bildNummer> + ".jpg"
	 *
	 * Beispiel:
	 * doc: "problem_schatzsuche_fehler_t1_3.psd" -> docPrefix = "problem_schatzsuche_fehler_"
	 * variantGroup: "t1_v6-10" -> variantPrefix = "t1_v", numbers 6..10
	 * imageGroup: "Bild 1" -> pSuffix = "_p1"
	 * -> "problem_schatzsuche_fehler_t1_v6_p1.jpg" ... "problem_schatzsuche_fehler_t1_v10_p1.jpg"
	 */

	if (app.documents.length === 0) {
		alert("Kein Dokument geöffnet.");
		return;
	}

	var doc = app.activeDocument;

	// =====================================================================
	// Farben ganz am Anfang definieren
	// =====================================================================

	/** Overlay-Farben */
	var COLOR_BLUE = [58, 110, 165];
	var COLOR_GOLD = [212, 178, 84];
	var COLOR_BEIGE = [232, 223, 201];
	var COLOR_CREME = [248, 244, 233];
	var COLOR_GREEN = [90, 138, 94];
	var COLOR_DARK_BLUE = [15, 45, 75];
	var COLOR_WHITE = [255, 255, 255];

	/**
	 * 5 Varianten: pro Variante [OverlayRGB, TextRGB].
	 */
	var VARIANT_COLORS = [
		[COLOR_BLUE, COLOR_WHITE],
		[COLOR_GOLD, COLOR_DARK_BLUE],
		[COLOR_BEIGE, COLOR_DARK_BLUE],
		[COLOR_GREEN, COLOR_CREME],
		[COLOR_DARK_BLUE, COLOR_BEIGE]
	];

	// =====================================================================
	// Konfiguration
	// =====================================================================

	var TOP_GROUP_IMAGES = "Bilder";
	var TOP_GROUP_VARIANTS = "Varianten";
	var OVERLAY_LAYER_NAME = "Overlay-Hintergrund";

	var JPG_QUALITY = 10; // 0..12
	var EXPORT_FOLDER = Folder("~/Downloads");

	// =====================================================================
	// Hilfsfunktionen
	// =====================================================================

	function getDocPrefix(fileName) {
		/**
		 * PSD-Namensschema:
		 * <Name>_t<zahl>.psd
		 * <Name>_t<zahl>_<zahl>.psd
		 * <Name>_t<zahl>_<zahl>_<zahl>.psd
		 * ...
		 *
		 * Ergebnis:
		 * <Name>_
		 */
		var base = fileName.replace(/\.[^\.]+$/, ""); // Extension weg

		// Schneidet alles ab "_t<zahl>" + optionale weitere "_<zahl>" am Ende
		// z.B.:
		// "problem_xyz_t2" -> "problem_xyz"
		// "problem_xyz_t2_3_4" -> "problem_xyz"
		var m = base.match(/^(.*)_t\d+(?:_\d+)*$/);
		if (m && m[1]) {
			return m[1] + "_";
		}

		// Fallback (falls Dateiname nicht dem Schema entspricht)
		return base + "_";
	}

	function parseVariantGroupName(groupName) {
		// Muster: t<zahl>_v<zahl>-<zahl>, z.B. "t1_v6-10"
		// -> prefix "t1_v", start 6, end 10
		var m = groupName.match(/^(t\d+_v)(\d+)\-(\d+)$/);
		if (!m) return null;

		return {
			prefix: m[1],
			start: parseInt(m[2], 10),
			end: parseInt(m[3], 10)
		};
	}

	function parseImageNumberFromGroupName(groupName) {
		// Erwartet: "Bild 1" (case-insensitive, beliebige Leerzeichen)
		// -> 1
		var m = groupName.match(/^\s*Bild\s+(\d+)\s*$/i);
		if (!m) return null;
		return parseInt(m[1], 10);
	}

	function makeSolidColor(rgbArr) {
		var c = new SolidColor();
		c.rgb.red = rgbArr[0];
		c.rgb.green = rgbArr[1];
		c.rgb.blue = rgbArr[2];
		return c;
	}

	function collectTextLayers(container, outArr) {
		var i;

		for (i = 0; i < container.artLayers.length; i++) {
			var l = container.artLayers[i];
			if (l.kind === LayerKind.TEXT) {
				outArr.push(l);
			}
		}

		for (i = 0; i < container.layerSets.length; i++) {
			collectTextLayers(container.layerSets[i], outArr);
		}
	}

	function findArtLayerByNameRecursive(container, targetName) {
		var i;

		for (i = 0; i < container.artLayers.length; i++) {
			if (container.artLayers[i].name === targetName) {
				return container.artLayers[i];
			}
		}

		for (i = 0; i < container.layerSets.length; i++) {
			var found = findArtLayerByNameRecursive(container.layerSets[i], targetName);
			if (found) return found;
		}

		return null;
	}

	function setTextLayersColor(textLayers, rgbArr) {
		var c = makeSolidColor(rgbArr);
		for (var i = 0; i < textLayers.length; i++) {
			try {
				textLayers[i].textItem.color = c;
			} catch (e) {
				// ignorieren
			}
		}
	}
	
	function selectLayerById(layerId) {
		var desc = new ActionDescriptor();
		var ref = new ActionReference();
		ref.putIdentifier(charIDToTypeID("Lyr "), layerId);
		desc.putReference(charIDToTypeID("null"), ref);
		executeAction(charIDToTypeID("slct"), desc, DialogModes.NO);
	}
	
	function applySolidFillColorViaListener(rgbArr) {
		/**
		 * Entspricht dem ScriptListener-Output:
		 * 1) Vordergrundfarbe setzen (FrgC)
		 * 2) contentLayer (Trgt) solidColorLayer Farbe setzen
		 *
		 * Erwartung: Die gewünschte SOLIDFILL-Formebene ist aktuell selektiert.
		 */
		var r = rgbArr[0];
		var g = rgbArr[1];
		var b = rgbArr[2];

		// --- 1) Foreground Color setzen ---
		var idsetd = charIDToTypeID("setd");
		var descFg = new ActionDescriptor();
		var idnull = charIDToTypeID("null");
		var refFg = new ActionReference();
		var idClr = charIDToTypeID("Clr ");
		var idFrgC = charIDToTypeID("FrgC");
		refFg.putProperty(idClr, idFrgC);
		descFg.putReference(idnull, refFg);

		var idT = charIDToTypeID("T   ");
		var descRgb1 = new ActionDescriptor();
		descRgb1.putDouble(charIDToTypeID("Rd  "), r);
		descRgb1.putDouble(charIDToTypeID("Grn "), g);
		descRgb1.putDouble(charIDToTypeID("Bl  "), b);
		descFg.putObject(idT, charIDToTypeID("RGBC"), descRgb1);

		descFg.putString(charIDToTypeID("Srce"), "photoshopPicker");
		executeAction(idsetd, descFg, DialogModes.NO);

		// --- 2) Solid Fill Farbe auf contentLayer (Target) setzen ---
		var descFill = new ActionDescriptor();
		var refFill = new ActionReference();
		refFill.putEnumerated(stringIDToTypeID("contentLayer"), charIDToTypeID("Ordn"), charIDToTypeID("Trgt"));
		descFill.putReference(idnull, refFill);

		var descSolid = new ActionDescriptor();
		var descClrOuter = new ActionDescriptor();
		var descRgb2 = new ActionDescriptor();
		descRgb2.putDouble(charIDToTypeID("Rd  "), r);
		descRgb2.putDouble(charIDToTypeID("Grn "), g);
		descRgb2.putDouble(charIDToTypeID("Bl  "), b);
		descClrOuter.putObject(charIDToTypeID("Clr "), charIDToTypeID("RGBC"), descRgb2);

		descFill.putObject(idT, stringIDToTypeID("solidColorLayer"), descClrOuter);

		executeAction(idsetd, descFill, DialogModes.NO);
	}

	function setOverlayColor(overlayLayer, rgbArr) {
		/**
		 * Selektiert die konkrete Overlay-Ebene und setzt dann die Farbe
		 * exakt wie Photoshop es beim Klick auf ein Bibliotheken-Farbfeld tut.
		 */
		selectLayerById(overlayLayer.id);
		applySolidFillColorViaListener(rgbArr);
	}

	function exportJpg(fullFile, quality) {
		var opts = new JPEGSaveOptions();
		opts.quality = quality;
		opts.embedColorProfile = true;
		opts.formatOptions = FormatOptions.STANDARDBASELINE;

		doc.saveAs(fullFile, opts, true, Extension.LOWERCASE);
	}

	function findTopLevelGroup(name) {
		for (var i = 0; i < doc.layerSets.length; i++) {
			if (doc.layerSets[i].name === name) {
				return doc.layerSets[i];
			}
		}
		return null;
	}

	function assertVariantCountMatchesRange(rangeInfo) {
		var expected = (rangeInfo.end - rangeInfo.start + 1);
		if (expected !== VARIANT_COLORS.length) {
			throw new Error(
				"Varianten-Gruppe '" + rangeInfo.prefix + rangeInfo.start + "-" + rangeInfo.end + "' erwartet " + expected + " Varianten,\n" +
				"aber VARIANT_COLORS hat " + VARIANT_COLORS.length + " Einträge.\n" +
				"Bitte VARIANT_COLORS oder die Range anpassen."
			);
		}
	}

	function hideAllChildGroups(parentGroup) {
		for (var i = 0; i < parentGroup.layerSets.length; i++) {
			parentGroup.layerSets[i].visible = false;
		}
	}
	
	function isLayerSetLocked(layerSet) {
		/**
		 * Erkennt, ob eine Gruppe (LayerSet) gesperrt ist.
		 * Je nach Photoshop-Version gibt es allLocked und/oder locked.
		 */
		try {
			if (layerSet.allLocked === true) {
				return true;
			}
		} catch (e) {
			// ignore
		}

		try {
			if (layerSet.locked === true) {
				return true;
			}
		} catch (e2) {
			// ignore
		}

		return false;
	}
	
	// =====================================================================
	// Hauptlogik
	// =====================================================================

	if (!EXPORT_FOLDER.exists) {
		EXPORT_FOLDER.create();
	}

	var imagesGroup = findTopLevelGroup(TOP_GROUP_IMAGES);
	if (!imagesGroup) {
		alert("Top-Level-Gruppe '" + TOP_GROUP_IMAGES + "' nicht gefunden.");
		return;
	}

	var variantsGroup = findTopLevelGroup(TOP_GROUP_VARIANTS);
	if (!variantsGroup) {
		alert("Top-Level-Gruppe '" + TOP_GROUP_VARIANTS + "' nicht gefunden.");
		return;
	}

	var docPrefix = getDocPrefix(doc.name);

	var originalRulerUnits = app.preferences.rulerUnits;
	app.preferences.rulerUnits = Units.PIXELS;

	var originalDialogMode = app.displayDialogs;
	app.displayDialogs = DialogModes.NO;

	try {
		// Beide Hauptgruppen sichtbar lassen, aber ihre Untergruppen steuern wir separat
		imagesGroup.visible = true;
		variantsGroup.visible = true;

		hideAllChildGroups(imagesGroup);
		hideAllChildGroups(variantsGroup);

		var exportedCount = 0;

		// 1) Jede Bilder-Untergruppe
		for (var b = 0; b < imagesGroup.layerSets.length; b++) {
			var imageSubGroup = imagesGroup.layerSets[b];
			
			if (isLayerSetLocked(imageSubGroup)) {
				continue;
			}

			var imageNumber = parseImageNumberFromGroupName(imageSubGroup.name);
			if (imageNumber === null) {
				// Falls du später andere Gruppen im Bilder-Ordner hast, werden die einfach übersprungen.
				continue;
			}

			// Nur diese Bilder-Untergruppe sichtbar
			hideAllChildGroups(imagesGroup);
			imageSubGroup.visible = true;

			var pSuffix = "_p" + imageNumber;

			// 2) Für diese Bilder-Untergruppe: alle Varianten-Gruppen
			for (var v = 0; v < variantsGroup.layerSets.length; v++) {
				var variantSubGroup = variantsGroup.layerSets[v];
				
				if (isLayerSetLocked(variantSubGroup)) {
					continue;
				}
				
				var rangeInfo = parseVariantGroupName(variantSubGroup.name);
				if (!rangeInfo) {
					continue;
				}

				assertVariantCountMatchesRange(rangeInfo);

				// Nur diese Varianten-Untergruppe sichtbar
				hideAllChildGroups(variantsGroup);
				variantSubGroup.visible = true;

				var overlay = findArtLayerByNameRecursive(variantSubGroup, OVERLAY_LAYER_NAME);
				if (!overlay) {
					throw new Error("In Varianten-Gruppe '" + variantSubGroup.name + "' wurde keine Ebene '" + OVERLAY_LAYER_NAME + "' gefunden.");
				}
				
				var textLayers = [];
				collectTextLayers(variantSubGroup, textLayers);

				// 3) Range exportieren (z.B. 6..10)
				for (var n = rangeInfo.start; n <= rangeInfo.end; n++) {
					var idx = (n - rangeInfo.start); // 0..4
					var overlayRgb = VARIANT_COLORS[idx][0];
					var textRgb = VARIANT_COLORS[idx][1];

					var fileName = docPrefix + rangeInfo.prefix + n + pSuffix + ".jpg";
					var fullFile = new File(EXPORT_FOLDER.fsName + "/" + fileName);

					var baseState = doc.activeHistoryState;

					try {
						setOverlayColor(overlay, overlayRgb);
						setTextLayersColor(textLayers, textRgb);

						exportJpg(fullFile, JPG_QUALITY);

					} finally {
						// Ganz wichtig: Nach JEDEM Export zurück auf den Ausgangszustand
						doc.activeHistoryState = baseState;
					}

					exportedCount++;
				}
			}
		}

		alert("Fertig. Exportierte Dateien: " + exportedCount + "\nZiel: " + EXPORT_FOLDER.fsName);

	} catch (e) {
		alert("Fehler:\n" + e.message);

	} finally {
		app.displayDialogs = originalDialogMode;
		app.preferences.rulerUnits = originalRulerUnits;
	}
})();
