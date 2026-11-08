Reconnaissance de texte (OCR) de Grimoire :
- tesseract.js 7.0.0 (Apache 2.0) : tesseract.min.js, worker.min.js
- tesseract.js-core 7.0.0 (Apache 2.0, moteur Tesseract compilé en WebAssembly), versions « LSTM » seulement : core/
- modèle français « 4.0.0_best_int » (tessdata, Apache 2.0, paquet @tesseract.js-data/fra) : lang/fra.traineddata.gz
Un seul des trois fichiers de core/ est téléchargé, selon ce que permet le navigateur.
