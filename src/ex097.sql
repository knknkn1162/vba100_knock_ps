SELECT
    T1.取引先CD,
    M1.取引先名,
    T1.商品CD,
    M2.商品名,
    SUM(T1.数量) AS 数量合計,
    SUM(T1.数量 * T1.単価) AS 金額合計,
    ROUND(SUM(T1.数量 * T1.単価) / SUM(T1.数量),0) AS 平均単価,
    M2.標準単価,
    S1.最低単価
FROM 
    ((([T売上] T1 LEFT JOIN [M取引先] AS M1 ON T1.取引先CD = M1.取引先CD)
        LEFT JOIN [M商品] AS M2 ON T1.商品CD = M2.商品CD)
            LEFT JOIN
                (SELECT 商品CD,MIN(単価) AS 最低単価 FROM [T売上] GROUP BY 商品CD)
                AS S1 ON T1.商品CD = S1.商品CD)
GROUP BY
    T1.取引先CD,M1.取引先名,T1.商品CD,M2.商品名,M2.標準単価,S1.最低単価
HAVING
    ROUND(SUM(T1.数量 * T1.単価) / SUM(T1.数量),0) > M2.標準単価
