(ns dave.datasource
  (:require [excellence.spreadsheet :as ex]))

(declare
 create-data-tree
 fetch-workbook-data
 inverse-map
 inverse-layer-2-3)

(defn create-data-tree [data-files]
  "Erzeugt den Datenbaum mit den Zellwerten aller Daten-Files.
Es werden die Daten aus den einzelnen Arbeitsmappen ausgelesen und dann
die erzeugte Baumstruktur umgebaut"
  (zipmap (keys data-files)
                  (for [wb-path-name (vals data-files)]
                    (-> (fetch-workbook-data wb-path-name)
                        inverse-map
                        inverse-layer-2-3))))

(defn fetch-workbook-data [wb-path-name]
  "Hilfsfunktion fuer 'create-data-tree'. Extrahiert alle Zellwerte aus
einer Arbeitsmappe. Es werden nur Sheets mit Jahreszahlen als Name einbezogen.
Erzeugter Datenbaum :blatt -> :zeile -> :spalte wert"
  (let [data-sheets (filter (fn [sheet] 
                              (not (nil? (re-find #"\d{4,4}" (ex/get-sheet-name sheet))))) 
                            (ex/sheet-seq (ex/workbook wb-path-name)))]
    (zipmap (map (fn [sh] (keyword (ex/get-sheet-name sh))) data-sheets)
            (for [sh data-sheets]
              (ex/indexed-value-map sh :coordinates-nested-map)))))

(defn- inverse-map [nested-map]
  "Hilfsfunktion fuer 'create-data-tree. Invertiert die erste 
mit zweiten Ebene einer 'map', z. B.
:jahr -> :zeile -> :spalte wert ==> :zeile -> :jahr -> spalte wert"
  (->> (map (fn [[outer innerS]]
              (map (fn [[inner value]]
                     [inner {outer value}])
                   innerS))
            nested-map)
       (reduce concat)
       (reduce (fn [res [inner entry]]
                 (update-in res [inner] merge entry))
               {})))

(defn- inverse-layer-2-3 [m]
  "Hilfsfunktion fuer 'create-data-tree' um die von 'inverse-map'
erzeugte Datenstruktur weiter anzupassen:
von   :zeile -> :jahr -> spalte wert
zu    :zeile -> :spalte -> jahr wert"
  (zipmap (keys m)
          (map inverse-map (vals m))))


