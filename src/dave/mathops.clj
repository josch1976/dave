(ns dave.mathops)

(defn- replace-with-0 [args]
  (map #(if (number? %) % 0) args))

(defn- apply-add-sub [args function]
  (let [arg-handler (fn [arg-list]
                      (if (every? (complement number?) arg-list)
                        nil
                        (apply function (replace-with-0 arg-list))))]
    (map arg-handler args)))

(defn add? [args]
  (apply-add-sub args +))

(defn sub? [args]
  (apply-add-sub args -))

(defn mul? [args]
  (let [arg-handler (fn [arg-list]
                      (if (some (complement number?) arg-list)
                            nil
                            (apply * arg-list)))]
    (map arg-handler args)))

(defn div? [args]
  (let [arg-handler (fn [arg-list]
                      (cond (every? (complement number?) arg-list)
                            nil
                            ((complement number?) (second arg-list))
                            nil
                            :else
                            (apply / (replace-with-0 arg-list))))]
    (map arg-handler args)))

(defn avg? [args]
  (/ (add? args) 
     (count (first args))))

(defn operation [operator args]
  (let [to-seq (map #(if (seq? %) % (list %)) args)
        max-count (reduce max (map count to-seq))
        equalized (map #(if (= 1 (count %))
                          (repeat max-count (first %))
                          %) 
                       to-seq)
        transposed (apply map list equalized)] 
    (operator transposed)))
