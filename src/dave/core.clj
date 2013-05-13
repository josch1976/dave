(ns dave.core
  (:require [excellence.spreadsheet :as ex])
  (:require [dave.logging :as dl])
  (:require [dave.datasource :as dd])
  (:require [dave.hooks :as dh] :reload :verbose)
  (:require [dave.mathops :as mo] :reload :verbose)
  (:gen-class))

(defn -main
  [& args])


;; Map zur Speicherung aller Daten-Dateien
(def data-files (atom {}))

;; der aktuell ausgewertete Monat
(def current-month (atom nil))

;; die darzustellenden Jahre
(def shown-years (atom nil))

;; Map aller Daten mit dem Aufbau :datei -> :zeile -> :spalte -> :jahr wert
;; pro Spalte kann es mehrere Jahre mit je einem Wert geben.
(def data-tree (atom {}))

;; Map mit Referenzen zu allen Zellen im Arbeitsblatt 'cfg=Konfiguration'
(def config-cells (atom {}))

;; Map aller im Arbeitsblatt 'cfg=Konfiguration' definierten Zellenformatierungen
(def cell-styles (atom {}))

;; Map aller im Arbeitsblatt 'cfg=Konfiguration' definierten Verweise auf Zellen
;; 'mit' Inhalt
(def cell-templates (atom {}))

(declare
 create-report-from-template
 ;; Konfiguration
 set-config-cells
 set-current-month
 set-shown-years
 create-configuration
 ;; Daten einlesen
 set-data-tree
 ;; Vorlagen-Arbeitsblaetter parsen
 process-template-sheets parse-eval-sheet read-eval-clojure-code
 try-read-string try-eval
 ;; Aufraeumarbeiten und Speichern
 delete-redundant-sheets save-created-report
 ;; Hooks fuer Funktionsaufrufe aus Excel
 add-data-file-alias add-cell-style-alias add-cell-template-alias
 insert-row-after-me 
 print-years get-vals get-list get-range print-vals print-trend print-list 
 print-achievement add? sub? mul? div?)

(defn do-main [& args]
"Argumente:
1. template-Datei
2. Monat
3. Jahre
4. logging-Ordner"
  (dl/config-logging (nth args 3))
  (dl/log! "INFO" "Anwendung gestartet")
  (set-current-month (nth args 1))
  (set-shown-years (nth args 2))
  (create-report-from-template (nth args 0)))

(comment (do-main 
          "/Users/anwender/Desktop/berichtswesen amt 60/gesamt.xls"
          "10"
          "2000, 2001, 2002"
          "/Users/anwender/dave-logging"))

(defn create-report-from-template [template-file]
  "Erzeuge den Report aus dem Template und speichere ihn unter einem
gewuenschten Pfad (Abfragemaske fuer Anwender"
  (let [wb (ex/load-workbook template-file)]
    (set-config-cells (ex/get-sheet wb "cfg=Konfiguration"))
    (create-configuration (ex/get-sheet wb "cfg=Konfiguration"))
    (set-data-tree)
    (process-template-sheets wb)
    (delete-redundant-sheets wb)
    (save-created-report wb)))

;; -----------------------------------------------------------------------------
;; Konfiguration
;; -----------------------------------------------------------------------------

(defn set-config-cells [config-sheet]
  (reset! config-cells (ex/indexed-cell-map config-sheet :address-key)))

(defn create-configuration [config-sheet]
  "Lade die Konfiguration fuer die Template-Erstellung
- 'map' mit den Daten-Files erzeugen
- 'map' mit den Zell-Formatierungen erzeugen"
  (dl/log! "INFO" "Lade Konfiguration fuer die Template-Erstellung")
  (doseq [c (ex/cell-seq config-sheet)]
    (read-eval-clojure-code (ex/get-cell-value c))))

(defn set-current-month [month]
  "Setzen der globalen Variable fuer den auszuwertenden Monat"
  (try
    (reset! current-month (Integer/parseInt month))
    (catch Exception e (dl/log! "FEHLER" (str "Monat '" 
                                             month 
                                             "' konnte nicht gesetzt werden")))))

(defn set-shown-years [chosen-years-string]
  "Setzen der globalen Variable fuer die anzuzeigenden Jahre"
  (reset! shown-years
          (doall (sort > (map (fn [y] (try	
                                 (Integer/parseInt (.trim y))
                                 (catch Exception 
                                     e 
                                   (dl/log! "FEHLER" (str "Anzuzeigende Jahre '" 
                                                          chosen-years-string 
                                                          "' konnten nicht gesetzt werden"))))) 
                       (.split chosen-years-string ","))))))

(defn set-data-tree []
  (dl/log! "INFO" "Daten-Baum erzeugen")
  (reset! data-tree (dd/create-data-tree @data-files)))


;; -----------------------------------------------------------------------------
;; Vorlagen-Arbeitsblaetter parsen
;; -----------------------------------------------------------------------------

(def current-cell (atom nil))

(defn process-template-sheets [wb]
  "Bearbeitet alle Arbeitsblaetter der Templatemappe aus denen Berichte erzeugt werden
sollen. Per Konvention beginnen diese immer mit 'aus='. Die Bearbeitung der einzelnen
Sheets wird an die Funktion modify-sheet weitergegeben."
  (let [templates (filter (fn [sh] (.startsWith (ex/get-sheet-name sh) "aus="))
                          (ex/sheet-seq wb))]
    (doseq [sh templates]
      (dl/log! "INFO" (str "Parse und evaluiere Arbeitsblatt '" (.getSheetName sh) "'"))
      (parse-eval-sheet sh))))

(defn parse-eval-sheet [sh]
  (reset! current-cell nil)
  (loop [cells (ex/cell-seq sh)]
    (when-let [cell (first cells)]
      (reset! current-cell cell)
      (let [v (ex/get-cell-value cell)]
        (if (and (string? v)
                 (.startsWith (.trim v) "<clj>"))
          (read-eval-clojure-code (subs (.trim v) 5))))
      (recur (next (drop-while (fn [tmp-cell] (not (= tmp-cell cell))) 
                               (ex/cell-seq sh)))))))

(defn try-read-string [string]
  (try (read-string string)
       (catch Exception 
           e 
         (do (dl/log! "FEHLER" (str "Blatt '" 
                                    (.getSheetName (.getSheet @current-cell)) "' "
                                    "Zeile ["(.getRowIndex @current-cell) "] " 
                                    "Spalte [" (.getColumnIndex @current-cell) "]"
                                    "\n"
                                    "read-string-Fehler: "
                                    "\n"
                                    string))
             (throw (Exception. "read-string-error"))))))

(defn try-eval [s-expression]
  (try (eval s-expression)
       (catch Exception 
           e
         (do (dl/log! "FEHLER" (str "Blatt '" 
                                    (.getSheetName (.getSheet @current-cell)) "' "
                                    "Zeile ["(.getRowIndex @current-cell) "] " 
                                    "Spalte [" (.getColumnIndex @current-cell) "]"
                                    "\n"
                                    "eval-string-Fehler: "
                                    "\n"
                                    s-expression))
             (throw (Exception. "eval-error"))))))

(defn read-eval-clojure-code [clj-string]
  (try
    (-> clj-string
        (try-read-string)
        (try-eval))
    (catch Exception e)))


;; -----------------------------------------------------------------------------
;; 'Aufraeumarbeiten' und Speicherung
;; -----------------------------------------------------------------------------

(def redundant-sheets 
  ["Einstellungen"
   "cfg=Konfiguration"])

(defn delete-redundant-sheets [wb]
"Loescht die fuer den Report nicht mehr benötigten Blaetter aus dem Template"
  (doseq [r-name redundant-sheets]
    (ex/delete-sheet! wb r-name)))

(defn save-created-report [wb]
  (ex/save-workbook! wb "/Users/anwender/Desktop/modifiziert.xls"))


;; -----------------------------------------------------------------------------
;; Hooks fuer Funktionsaufrufe aus Excel
;; -----------------------------------------------------------------------------

(defn add-data-file-alias [alias-key file-path]
  (dh/add-data-file-alias data-files alias-key file-path))

(defn add-cell-style-alias [alias-key address-key]
  (dh/add-cell-style-alias alias-key address-key config-cells cell-styles))

(defn add-cell-template-alias [alias-key address-key]
  (dh/add-cell-template-alias alias-key address-key config-cells cell-templates))

(defn insert-row-after-me []
  (dh/insert-row-after-me current-cell))

(defn print-years [& args] 
  (apply dh/print-years (concat [:current-cell current-cell
                                 :shown-years shown-years 
                                 :cell-styles cell-styles] 
                                args)))

(defn get-vals [& args]
  (apply dh/get-vals (concat [data-tree shown-years] args)))

(defn get-list [& args]
  (dh/get-vals args))

(defn get-range [& args]
  (dh/get-range args))

(defn print-vals [& args]
  (apply dh/print-vals (concat [:current-cell current-cell 
                                :cell-styles cell-styles] 
                               args)))
(defn print-trend [& args]
  (apply dh/print-trend (concat [:current-cell current-cell
                                 :current-month current-month
                                 :cell-templates cell-templates]
                                args)))

(defn print-achievement [& args]
  (apply dh/print-achievement (concat [:current-cell current-cell
                                       :cell-templates cell-templates]
                                      args)))

(defn printl-list [& args]
  (apply dh/print-list (concat [:current-cell current-cell
                                :cell-styles cell-styles]
                               args)))

(defn add? [& args]
  (mo/operation mo/add? args))

(defn sub? [& args]
  (mo/operation mo/sub? args))

(defn mul? [& args]
  (mo/operation mo/mul? args))

(defn div? [& args]
  (mo/operation mo/div? args))