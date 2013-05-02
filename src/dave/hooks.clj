(ns dave.hooks
  (:require [excellence.spreadsheet :as ex]))

(declare 
 add-data-file-alias
 insert-row-after-me
 print-years
 get-vals
 get-list
 get-range
 print-vals
 print-trend
 print-achievement
 print-list)


(defn add-data-file-alias [data-files alias-key file-path]
  (swap! data-files assoc alias-key file-path))

(defn insert-row-after-me [current-cell]
  (ex/insert-row-after! (.getRow @current-cell)))

(defn add-cell-style-alias [alias-key address-key config-cells cell-styles]
  (swap! cell-styles
         assoc 
         alias-key 
         (.getCellStyle (address-key @config-cells))))

(defn add-cell-template-alias [alias-key address-key config-cells cell-templates]
  (swap! cell-templates
         assoc 
         alias-key 
         (address-key @config-cells)))

(defn do-offset [current-cell times]
  (dotimes [_ times]
    (ex/insert-row-after! (.getRow @current-cell))))

(defn get-next [cell orientation]
  (if (= orientation :down) 
    (ex/create-cell! (ex/get-row (.getSheet cell) (+ 1 (.getRowIndex cell)))
                     (.getColumnIndex cell))
    (ex/create-cell! (.getRow cell)
                     (+ 1 (.getColumnIndex cell)))))

(defn print-current-previous
  [print-values cell orientation offset styles style-current show-previous style-previous]
                                        ; aktuelles Jahr
  (ex/set-cell-value! cell (first print-values))
  (.setCellStyle cell (style-current styles))
                                        ; vorige Jahre
  (when show-previous
    (loop [cell cell
           previous (rest print-values)]
      (when (seq previous)
        (when offset (ex/insert-row-after! (.getRow cell)))
        (let [new-cell (get-next cell orientation)]
          (ex/set-cell-value! new-cell (first previous))
          (.setCellStyle new-cell (get styles 
                                       (if (nil? style-previous) 
                                         style-current 
                                         style-previous)))
          (recur new-cell (rest previous)))))))

(defn print-years [& {:keys [current-cell
                             offset
                             shown-years
                             cell-styles
                             style-current
                             show-previous
                             style-previous
                             orientation]
                      :or {show-previous true
                           style-previous nil
                           offset nil
                           orientation :down}}]
  (let [print-values (map str @shown-years)
        cell @current-cell
        styles @cell-styles]
    (print-current-previous print-values
                            cell
                            orientation
                            offset
                            styles
                            style-current
                            show-previous
                            style-previous)))

(defn get-vals [data-tree shown-years & args]
  (let [keywordized (map (comp keyword str) @shown-years)]
    (map (get-in @data-tree args) keywordized)))

(defn get-list [args]
  (let [del-nil (fn [c] (if (nil? c) "-" c))] 
    (map (comp del-nil first) args)))

(defn get-range [args]
  )

(defn print-vals [& {:keys [current-cell
                            offset
                            values
                            cell-styles
                            style-current
                            show-previous
                            style-previous
                            orientation]
                     :or {offset nil
                          show-previous true
                          style-previous nil
                          orientation :down}}]
  (let [print-values (map #(if (nil? %) "-" %) values)
        cell @current-cell
        styles @cell-styles]
    (print-current-previous print-values
                            cell
                            orientation
                            offset
                            styles
                            style-current
                            show-previous
                            style-previous)))

(defn project-on-year [value months]
  (if (nil? value) 
    nil
    (* (/ value 12) months)))

(defn print-trend [& {:keys [current-cell
                             current-month
                             cell-templates
                             values]}]
  (let [to-examine (project-on-year (first values) @current-month)
        trend (cond (nil? to-examine) :-?- 
                    (<= to-examine -15) :x<=-15
                     (and (> to-examine -15)
                          (< to-examine -5)) :-15>x<-5
                     (and (>= to-examine -5)
                          (<= to-examine 5)) :-5>=x<=5
                     (and (> to-examine 5)
                          (< to-examine 15)) :5>x<15
                     (>= to-examine 15) :x>=15
                     :else :-?-)
        template (get @cell-templates trend)
        cell @current-cell]
    (ex/set-cell-value! cell (ex/get-cell-value template))
    (.setCellStyle cell (.getCellStyle template))))

(defn print-achievement [& {:keys [current-cell
                                   cell-templates
                                   values]}]
  (let [to-examine (first values)
        achievement (cond (nil? to-examine) :-?- 
                          (< to-examine 0) :x<0
                          :else :x>=0)
        template (get @cell-templates achievement)
        cell @current-cell]
    (ex/set-cell-value! cell (ex/get-cell-value template))
    (.setCellStyle cell (.getCellStyle template))))

(defn print-list [[& {:keys [current-cell
                             offset
                             cell-styles
                             style]
                      :or {offset nil}}]]
  )