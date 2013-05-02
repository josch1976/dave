(ns dave.logging
  (:import [java.io File])
  (:import [java.util Date])
  (:import [java.text SimpleDateFormat]))

(declare
 config-logging
 log!)

;; Speicherort fuer die Log-Dateien
(def log-folder (atom nil))

;; Generische Erzeugung eines Dateinamens fuer die Log-Datei
(def log-file (str "dave-log_" 
                   (.format (SimpleDateFormat. "yyyy-MM-dd__HH-mm--ss-SSS") 
                       (doto (Date.) 
                         (.setTime (System/currentTimeMillis))))
                   ".txt"))

(defn config-logging [folder]
"Einrichtung des Ordners wo die Log-Dateien gespeichert werden."
  (if (.exists (File. folder))
      (reset! log-folder folder)
      (do
        (let [default-folder "/Users/anwender/Desktop/dave_logging/"]
          (when-not (.exists (File. default-folder))
            (.mkdir (File. default-folder)))
          (reset! log-folder default-folder)))))

(defn log! [level message & stacktrace]
  "Schreibt eine Log-Nachricht in die Log-Datei"
  (spit (str @log-folder "/" log-file)  
        (str level "\n"
             message "\n"
             (when (> (count stacktrace) 0) (str (first stacktrace) "\n"))
             "\n\n")
        :append true))