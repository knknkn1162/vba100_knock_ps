SELECT 
    T1.取引先CD,
    M1.取引先名,
    T1.商品CD,
    M2.商品名,
    T1.日付,
    T1.単価,
    T1.数量,
    T1.数量 * T1.単価 AS 金額
 FROM 
    (([T売上] As T1 LEFT JOIN [M取引先] AS M1 ON T1.取引先CD = M1.取引先CD)
            LEFT JOIN [M商品] AS M2 ON T1.商品CD = M2.商品CD)
 WHERE
    T1.日付 >= #$date# 
        AND T1.数量 * T1.単価 >= $price
