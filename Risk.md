JGB Trader Model VaR Management

1. 目的

本ドキュメントは、JGBポートフォリオを年限別グリッドからトレーダーモデルへ変換し、以下をExcelで計算・管理するための理論と実装を整理したものである。

• 分散共分散法によるVaR
• VaR Allocation（Euler／Aumann–Shapley配分）
• Marginal VaRおよびComponent VaR
• 単独VaRとの違い
• 単一商品・複数商品によるVaR最小化
• VaR最小化後の変化額 (\Delta c)
• 年限モデルからトレーダーモデルへの共分散変換
• Excel VBAによるシート関数

本書で用いられる「CoVaR」は、共分散行列を使ったVariance–Covariance VaRを指す。システミックリスク研究のCoVaRや、Expected Shortfallの別名として使われるCVaRとは別概念である。

────────

2. 記号

|記号                         |意味                   |
|---------------------------|---------------------|
|(S_{\mathrm{grid}})        |年限別BPVベクトル           |
|(S_{\mathrm{trade}})       |トレーダーモデル上のリスクベクトル    |
|(J)                        |TradeMatrix          |
|(\Delta r_{\mathrm{grid}}) |年限別金利変化              |
|(\Delta r_{\mathrm{trade}})|Base Trade別の市場変化     |
|(Q_{\mathrm{grid}})        |年限別金利変化の共分散行列        |
|(Q_{\mathrm{trade}})       |トレーダーモデル上の共分散行列      |
|(c)                        |信頼係数を掛ける前の1標準偏差PnLリスク|
|(z)                        |信頼係数                 |
|(V)                        |信頼係数を掛けたVaR          |
|(e_i)                      |第(i)成分のみ1の単位ベクトル     |

典型的な単位は以下のとおりである。

|項目  |単位  |
|----|----|
|BPV |円/bp|
|金利変化|bp  |
|共分散 |bp² |
|VaR |円   |

────────

3. 年限モデルからトレーダーモデルへの変換

3.1 リスクベクトル

TradeMatrixを次のように定義する。

[
S_{\mathrm{grid}}=J S_{\mathrm{trade}}
]

(J)が正方かつフルランクなら、

[
\boxed{
S_{\mathrm{trade}}=J^{-1}S_{\mathrm{grid}}
}
]

となる。

各列は、例えば次のようなBase Tradeを表す。

[
2s5s=-2Y+5Y
]

[
2s3s4s=+2Y-2\times3Y+4Y
]

トレーダーモデルの各成分は、実際の銘柄保有額ではなく、年限別BPVをBase Tradeの組み合わせで再現した座標係数である。

3.2 市場変化

Base Tradeの市場変化は、

[ \boxed{ \Delta r_{\mathrm{trade}}

J^T\Delta r_{\mathrm{grid}}
}
]

である。

3.3 共分散行列

したがって、トレーダーモデル上の共分散行列は、

[ \boxed{ Q_{\mathrm{trade}}

J^TQ_{\mathrm{grid}}J
}
]

となる。

リスクベクトルと共分散行列を整合的に変換すれば、

[ S_{\mathrm{trade}}^TQ_{\mathrm{trade}}S_{\mathrm{trade}}

S_{\mathrm{grid}}^TQ_{\mathrm{grid}}S_{\mathrm{grid}}
]

となり、ポートフォリオ全体のVaRは座標変換によって変わらない。

> 注意  
> (Q_{\mathrm{trade}}=J^{-1}Q_{\mathrm{grid}}(J^{-1})^T)ではない。  
> 本モデルの定義では、正しい式は (Q_{\mathrm{trade}}=J^TQ_{\mathrm{grid}}J) である。

────────

4. VaRの計算

ポートフォリオの線形PnLを、

[
\Delta P=S^T\Delta r
]

とする。

PnL分散は、

[
\operatorname{Var}(\Delta P)=S^TQS
]

したがって、1標準偏差のPnLリスクは、

[
\boxed{
c=\sqrt{S^TQS}
}
]

信頼係数(z)を掛けたVaRは、

[
\boxed{
V=z\sqrt{S^TQS}
}
]

である。

代表的な片側信頼係数は以下のとおり。

|信頼水準|(z)  |
|----|----:|
|95% |1.645|
|99% |2.326|

────────

5. VaR Allocation

VaR Allocationは、ポートフォリオ全体のVaRを、各リスクファクターまたはBase Tradeへ配分する方法である。

5.1 Marginal VaR

[
V=z\sqrt{S^TQS}
]

を(S_i)で偏微分すると、

[ \boxed{ \mathrm{MVaR}_i

\frac{\partial V}{\partial S_i}

z\frac{(QS)_i}{\sqrt{S^TQS}}
}
]

となる。

これは、Trade (i) のリスクを1単位追加した場合に、全体VaRが限界的にどれだけ変化するかを表す。

5.2 Component VaR

[ \boxed{ \mathrm{CVaR}_i

S_i\mathrm{MVaR}_i

z\frac{S_i(QS)_i}{\sqrt{S^TQS}}
}
]

ベクトル形式では、

[ \boxed{ \mathrm{CVaR}

z\frac{\operatorname{diag}(S)QS}{\sqrt{S^TQS}}
}
]

である。

Euler配分の性質により、

[
\boxed{
\sum_i \mathrm{CVaR}_i=V
}
]

が成立する。

5.3 VaR Share

各TradeのVaR寄与率は、

[ \boxed{ \mathrm{VaRShare}_i

\frac{\mathrm{CVaR}_i}{V}
}
]

である。

合計は100%になる。ただし、個別値は負になったり、100%を超えたりすることがある。

5.4 負のComponent VaR

Component VaRが負の場合、そのTradeは現在のポートフォリオ内でヘッジとして機能している。

したがって、

• BPVが大きい
• 単独VaRが大きい
• Component VaRが大きい

はそれぞれ異なる意味を持つ。

負のComponent VaRを持つポジションを機械的に閉じると、全体VaRが増加する場合がある。

────────

6. Standalone VaRとの違い

Trade (i)だけを単独で保有した場合のVaRは、

[ \boxed{ \mathrm{StandaloneVaR}_i

z|S_i|\sqrt{q_{ii}}
}
]

である。

一方、Component VaRは、

[ \boxed{ \mathrm{CVaR}_i

z\frac{S_i(QS)_i}{\sqrt{S^TQS}}
}
]

である。

6.1 Standalone VaR

答える質問：

> このTradeだけを保有した場合、どれだけのリスクがあるか。

特徴：

• 常に0以上
• 他のTradeとの相関を無視
• 各Trade自身のポジション量とボラティリティのみ反映

6.2 Component VaR

答える質問：

> 現在のポートフォリオ全体の中で、このTradeがVaRにどれだけ寄与しているか。

特徴：

• 正にも負にもなる
• 他のTradeとの相関を反映
• 合計がポートフォリオVaRになる

6.3 Close-out VaR Change

Trade (i)を全量閉じた場合の実際のVaR変化は、

[ \boxed{ \Delta V_i^{\mathrm{close}}

V(S-S_ie_i)-V(S)
}
]

である。

Component VaRは限界感応度に基づく加法的配分であり、全量閉鎖時の有限変化とは一般に一致しない。

────────

7. 単一商品によるVaR最小化

商品 (i)だけを(x)取引する場合、新しいリスクは、

[
S^{\mathrm{new}}=S+xe_i
]

となる。

取引後の分散は、

[
(S+xe_i)^TQ(S+xe_i)
]

これを(x)について最小化すると、

[ \boxed{ x_i^*

-\frac{(QS)i}{q{ii}}
}
]

となる。

すべての商品について計算したベクトルは、

[ \boxed{ S^{\mathrm{trade}}

-\operatorname{diag}(\operatorname{diag}Q)^{-1}QS
}
]

である。

ただし、各成分は「その商品だけを取引する場合」の独立した候補であり、全成分を同時に実行してはならない。

────────

8. 図中の (\Delta c) の計算

8.1 取引前

[ \boxed{ c_{\mathrm{before}}

\sqrt{S^TQS}
}
]

8.2 任意の取引量(x)

商品 (i)を(x)取引した後の1σリスクは、

[ c_{\mathrm{new}}(x)

\sqrt{(S+xe_i)^TQ(S+xe_i)}
]

展開すると、

[ \boxed{ c_{\mathrm{new}}(x)

\sqrt{
c_{\mathrm{before}}^2
+
2x(QS)i
+
x^2q{ii}
}
}
]

したがって、符号付き変化額は、

[ \boxed{ \Delta c(x)

c_{\mathrm{new}}(x)-c_{\mathrm{before}}
}
]

である。

• (\Delta c<0)：VaR減少
• (\Delta c>0)：VaR増加

8.3 最適取引量を使った場合

[ x_i^*

-\frac{(QS)i}{q{ii}}
]

を代入すると、

[ \boxed{ c_i^{\min}

\sqrt{
c^2-\frac{(QS)i^2}{q{ii}}
}
}
]

よって、

[ \boxed{ \Delta c_i

c_i^{\min}-c

\sqrt{
c^2-\frac{(QS)i^2}{q{ii}}
}
-c
}
]

となる。

VaR削減額を正の値で表示する場合は、

[ \boxed{ \mathrm{VaRReduction}_i

c-c_i^{\min}
}
]

VaR削減率は、

[ \boxed{ \mathrm{ReductionRate}_i

\frac{c-c_i^{\min}}{c}
}
]

である。

信頼係数(z)を掛ける場合、取引前後の両方へ同じ(z)を掛ける。

────────

9. 複数商品によるVaR最小化

取引可能な複数商品のリスク方向を列に持つ行列を(H)とする。

[
S^{\mathrm{new}}=S+Hx
]

このとき、VaRを最小化する取引量は、

[ \boxed{ x^*

-(H^TQH)^{-1}H^TQS
}
]

となる。

取引後リスクは、

[
\boxed{
S^{\mathrm{new}}=S+Hx^*
}
]

である。

9.1 トレーダーモデルでの(H)

• トレーダーモデル座標上で特定のBase Tradeを選ぶ場合：(H)は選択行列
• 年限グリッド上でBase Tradeを取引する場合：(H)はTradeMatrix (J) の選択列

9.2 注意事項

• (H^TQH)が特異または悪条件の場合、逆行列が不安定になる
• 線形従属するトレードを同時に選ばない
• 実務では取引コスト、流動性、数量上限を考慮する
• 全ファクターを無制約で取引できる場合、理論上はリスクをゼロにできてしまうため、実務上の制約が必要

────────

10. 分散効果

共分散ベースのリスク尺度、

[
c(S)=\sqrt{S^TQS}
]

は、(Q)が半正定値なら劣加法性を満たす。

[
\boxed{
c(S_A+S_B)
\le
c(S_A)+c(S_B)
}
]

分散効果は、

[ \boxed{ \mathrm{DiversificationBenefit}

c(S_A)+c(S_B)-c(S_A+S_B)
}
]

で測定できる。

────────

11. Excel VBA実装

以下を標準モジュールに貼り付ける。

```vb
Option Explicit

'=========================================================
' VaRCalc
'
' VaR = Z × Sqr(S'QS)
'
' 使用例:
' =VaRCalc(B2:B16,D2:R16)
' =VaRCalc(B2:B16,D2:R16,2.326)
'=========================================================
Public Function VaRCalc( _
    ByVal RiskVector As Range, _
    ByVal CovMatrix As Range, _
    Optional ByVal Z As Double = 1# _
) As Variant

    Dim S As Variant
    Dim QS As Variant
    Dim Variance As Double
    Dim n As Long

    On Error GoTo ErrorHandler

    S = ToColumnVector(RiskVector)
    n = UBound(S, 1)

    If CovMatrix.Rows.Count <> n _
       Or CovMatrix.Columns.Count <> n Then
        VaRCalc = CVErr(xlErrRef)
        Exit Function
    End If

    If Z < 0# Then
        VaRCalc = CVErr(xlErrNum)
        Exit Function
    End If

    QS = Application.MMult(CovMatrix.Value2, S)
    Variance = CDbl(Application.SumProduct(S, QS))

    If Variance < -0.0000000001 Then
        VaRCalc = CVErr(xlErrNum)
        Exit Function
    End If

    Variance = Application.Max(0#, Variance)
    VaRCalc = Z * Sqr(Variance)
    Exit Function

ErrorHandler:
    VaRCalc = CVErr(xlErrValue)

End Function


'=========================================================
' VaRAllocate
'
' Allocation_i
' = Z × S_i × (QS)_i / Sqr(S'QS)
'
' 使用例:
' =VaRAllocate(B2:B16,D2:R16)
' =VaRAllocate(B2:B16,D2:R16,2.326)
'=========================================================
Public Function VaRAllocate( _
    ByVal RiskVector As Range, _
    ByVal CovMatrix As Range, _
    Optional ByVal Z As Double = 1# _
) As Variant

    Dim S As Variant
    Dim QS As Variant
    Dim Result() As Double
    Dim Variance As Double
    Dim BaseRisk As Double
    Dim n As Long
    Dim i As Long

    On Error GoTo ErrorHandler

    S = ToColumnVector(RiskVector)
    n = UBound(S, 1)

    If CovMatrix.Rows.Count <> n _
       Or CovMatrix.Columns.Count <> n Then
        VaRAllocate = CVErr(xlErrRef)
        Exit Function
    End If

    If Z < 0# Then
        VaRAllocate = CVErr(xlErrNum)
        Exit Function
    End If

    QS = Application.MMult(CovMatrix.Value2, S)
    Variance = CDbl(Application.SumProduct(S, QS))

    If Variance < -0.0000000001 Then
        VaRAllocate = CVErr(xlErrNum)
        Exit Function
    End If

    Variance = Application.Max(0#, Variance)
    BaseRisk = Sqr(Variance)

    ReDim Result(1 To n, 1 To 1)

    If BaseRisk = 0# Then
        VaRAllocate = Result
        Exit Function
    End If

    For i = 1 To n
        Result(i, 1) = _
            Z * CDbl(S(i, 1)) * CDbl(QS(i, 1)) / BaseRisk
    Next i

    If RiskVector.Rows.Count = 1 Then
        VaRAllocate = Application.Transpose(Result)
    Else
        VaRAllocate = Result
    End If

    Exit Function

ErrorHandler:
    VaRAllocate = CVErr(xlErrValue)

End Function


'=========================================================
' MatrixTBA
'
' A' × B × A を返す
'
' A : m行×n列
' B : m行×m列
' 出力 : n行×n列
'
' 使用例:
' =MatrixTBA(TradeMatrix,CovGrid)
'=========================================================
Public Function MatrixTBA( _
    ByVal A As Range, _
    ByVal B As Range _
) As Variant

    On Error GoTo ErrorHandler

    If B.Rows.Count <> B.Columns.Count Then
        MatrixTBA = CVErr(xlErrValue)
        Exit Function
    End If

    If B.Rows.Count <> A.Rows.Count Then
        MatrixTBA = CVErr(xlErrRef)
        Exit Function
    End If

    MatrixTBA = Application.MMult( _
        Application.Transpose(A.Value2), _
        Application.MMult(B.Value2, A.Value2) _
    )

    Exit Function

ErrorHandler:
    MatrixTBA = CVErr(xlErrValue)

End Function


'=========================================================
' 縦・横ベクトルを縦ベクトルへ統一
'=========================================================
Private Function ToColumnVector( _
    ByVal InputRange As Range _
) As Variant

    If InputRange.Columns.Count = 1 Then
        ToColumnVector = InputRange.Value2

    ElseIf InputRange.Rows.Count = 1 Then
        ToColumnVector = Application.Transpose(InputRange.Value2)

    Else
        Err.Raise vbObjectError + 1000, _
                  "ToColumnVector", _
                  "RiskVectorは1列または1行で指定してください。"
    End If

End Function
```

────────

12. 使用例

トレーダーモデルBPVをB2:B16、トレーダーモデル共分散行列をD2:R16とする。

1標準偏差リスク

```excel
=VaRCalc(B2:B16,D2:R16)
```

99% VaR

```excel
=VaRCalc(B2:B16,D2:R16,2.326)
```

99% VaR Allocation

```excel
=VaRAllocate(B2:B16,D2:R16,2.326)
```

Allocationの検算

```excel
=SUM(VaRAllocate(B2:B16,D2:R16,2.326))
```

は、

```excel
=VaRCalc(B2:B16,D2:R16,2.326)
```

と一致する。

トレーダーモデル共分散行列

```excel
=MatrixTBA(TradeMatrix,CovGrid)
```

は、

[ Q_{\mathrm{trade}}

J^TQ_{\mathrm{grid}}J
]

を返す。

────────

13. 推奨表示項目

|Base Trade|Current BPV|Absolute BPV|Standalone VaR|Marginal VaR|Component VaR|VaR Share|Single Optimal Trade|(\Delta c)|Reduction Rate|
|----------|----------:|-----------:|-------------:|-----------:|------------:|--------:|-------------------:|---------:|-------------:|

各列の意味

• Current BPV：トレーダーモデル上の符号付き感応度
• Absolute BPV：1bp感応度の絶対量
• Standalone VaR：そのTradeだけを保有した場合のVaR
• Marginal VaR：リスクを1単位追加した場合のVaR感応度
• Component VaR：現在の全体VaRへの寄与
• VaR Share：Component VaRを全体VaRで割った割合
• Single Optimal Trade：そのTradeだけでVaRを最小化する取引量
• (\Delta c)：最適化前後のVaR変化
• Reduction Rate：最適化によるVaR削減率

────────

14. 実務上の留意点

1. 共分散行列は金利水準ではなく、原則として日次金利変化から計算する。
2. BPVと金利変化の単位を必ず一致させる。
3. TradeMatrixの符号、スプレッド定義、フライの正規化を固定する。
4. トレーダーモデルのVaR Allocationは、選択した座標系に依存する。
5. トレーダーモデルをVaR配分に使う場合、TradeMatrixは原則としてフルランクの固定基底とする。
6. Component VaRが負でも計算エラーとは限らない。
7. Component VaRは全量閉鎖時のVaR変化ではない。
8. 単一商品最適化の候補をすべて同時に執行してはならない。
9. 複数商品最適化では、流動性・取引コスト・数量上限を別途考慮する。
10. VaRだけでは、銘柄固有リスク、レポ特殊性、入札、日銀オペ、CTD、流動性ショックを十分に捉えられないため、ストレステスト等と併用する。

────────

15. 核心

本モデルの目的は、VaRを単なる報告用の数字ではなく、

• どのBase Tradeがリスクを作っているか
• どのポジションがヘッジとして機能しているか
• 何をどれだけ取引すればVaRが低下するか
• 取引後にVaRがどれだけ変化するか

を判断するための、トレーダー向けポートフォリオ管理ツールとして利用することである。