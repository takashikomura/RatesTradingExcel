# PCA・TradeMatrix・VaRによる統合金利リスク管理フレームワーク

## 1. 目的と基本思想

本ドキュメントでは、JGBポートフォリオの金利リスクを、Principal Component Analysis（PCA）、TradeMatrix、およびValue at Risk（VaR）を用いて体系的に管理するための数学的フレームワークを整理する。

従来の金利リスク管理では、以下の3つの分析が独立して行われることが多い。

1. **Tenor Risk Management**：年限別のBPV・Key Rate Riskによる管理
2. **PCA Risk Management**：Level・Slope・Curvatureなどの統計的リスクファクターによる管理
3. **Trader Model**：Outright・Curve・Butterflyなど、トレーダーが認識するトレード構造による管理

本フレームワークの目的は、これらを一つの線形代数的な体系に統合することである。

特に、TradeMatrixのBase TradeをPCA Hedge Ratioによって定義することで、トレーダーが認識するポジション構造と、PCAによるリスクファクター分析を整合させる。

最終的には、以下を実現する。

- 年限別BPVをOutright・Curve・Flyに一意に分解する。
- 各Base Tradeが保有するPCA Factor Exposureを計算する。
- ポートフォリオ全体のVaRを、年限別・Trade別・PC別に分解する。
- PCとBase Tradeの二次元VaR Attribution Matrixを構築する。
- Signed ExposureとVaR Shockによって、トレードの方向性とリスク寄与の関係を理解する。

**重要な設計原則：PCAはリスク構造を定義・測定するためのモデルであり、トレードの期待収益や平均回帰を予測するモデルではない。**

---

## 2. 数学的枠組み：3つの空間

本フレームワークでは、ポートフォリオを3つの異なる空間で表現する。

### 2.1 Tenor Space

年限ごとの金利リスクを表す空間。

```math
b\in\mathbb{R}^{n}
```

```math
\Delta y\in\mathbb{R}^{n}
```

```math
\Sigma=\mathrm{Cov}(\Delta y)\in\mathbb{R}^{n\times n}
```

ここで、

- `b`：Signed BPVベクトル
- `Δy`：年限別金利変化ベクトル
- `Σ`：金利変化の共分散行列

である。

BPVは通貨額/bp、金利変化はbp単位とする。

ポートフォリオの一次近似P&Lは、

```math
\boxed{
\Delta PL=-b^\top\Delta y
}
```

で与えられる。

ここでは、正のBPVを金利上昇によって損失が発生するLong Duration Exposureとして定義する。

### 2.2 Trade Space

Outright・Curve・FlyなどのBase Tradeによってポートフォリオを表現する空間。

```math
x\in\mathbb{R}^{n}
```

```math
T\in\mathbb{R}^{n\times n}
```

ここで、

- `T`：TradeMatrix
- `x`：各Base TradeのExposure

である。

```math
\boxed{
b=Tx
}
```

TradeMatrixが正則であれば、

```math
\boxed{
x=T^{-1}b
}
```

となる。

### 2.3 PCA Space

金利変化の共分散構造を、直交するリスクファクターで表現する空間。

```math
\Sigma=V\Lambda V^\top
```

ここで、

- `V`：固有ベクトル行列
- `Λ`：固有値対角行列

である。

```math
V^\top V=I
```

PCA Factor Exposureを、

```math
\boxed{
q=V^\top b
}
```

と定義する。

### 2.4 全体の関係

```text
                  Tenor Space
              ┌────────────────┐
              │                │
              │  BPV       b   │
              │  Yield     Δy  │
              │  Cov       Σ   │
              │                │
              └────────────────┘
                    │      │
             T^(-1) │      │ V^T
                    ▼      ▼
           Trade Space    PCA Space
         ┌─────────────┐ ┌─────────────┐
         │ x = T^(-1)b │ │ q = V^T b   │
         │ Δs = T^T Δy │ │ Δf = V^T Δy │
         │ ΣT = T^TΣT  │ │ Q = V^TΣV   │
         └─────────────┘ └─────────────┘
                  \          /
                   \        /
                    ▼      ▼
                Integrated Risk
                 Attribution
```

3つの空間は異なるポートフォリオを表しているわけではない。同一のポートフォリオを異なる基底で表現している。

---

## 3. TradeMatrixの数学的性質

### 3.1 TradeMatrixの定義

Base Tradeを構成するSigned BPVウェイトベクトルを、

```math
t_1,t_2,\ldots,t_n\in\mathbb{R}^{n}
```

とする。

TradeMatrixは、

```math
\boxed{
T=
\begin{pmatrix}
t_1&t_2&\cdots&t_n
\end{pmatrix}
}
```

である。

各列が1つのBase Tradeを表し、各行が1つのTenorを表す。

任意のポートフォリオは、

```math
\boxed{
b=\sum_{k=1}^{n}x_kt_k=Tx
}
```

と表現される。

この分解における `x` は、実際に執行されたトレードの履歴ではなく、**選択したBase Tradeを基底とする合成的なポジション表現**である。

### 3.2 フルランク性

任意のBPVベクトルを一意に分解するための必要十分条件は、

```math
\boxed{
\mathrm{rank}(T)=n
}
```

である。

この条件が満たされる場合、

```math
x=T^{-1}b
```

が一意に決まる。

フルランク性はBase Tradeの線形独立性を保証するが、Base Trade同士が直交していることは要求しない。

### 3.3 条件数

TradeMatrixの数値的安定性は条件数によって評価できる。

```math
\boxed{
\kappa_2(T)=
\frac{\sigma_{\max}(T)}
{\sigma_{\min}(T)}
}
```

条件数が大きいほど、BPVの小さな変化がTrade Exposureの大きな変化につながり得る。

ただし、条件数は列のスケールに依存する。

Base Trade間の近接性を評価する場合には、各列をユークリッドノルムで正規化したTradeMatrixの条件数も確認する。

```math
\widetilde T=TD^{-1}
```

```math
D=\mathrm{diag}
\left(
\|t_1\|_2,\ldots,\|t_n\|_2
\right)
```

フルランクであることと、数値的に安定していることは別の性質である。

---

## 4. 採用するJGB TradeMatrix

### 4.1 対象Tenor

今回のモデルでは、以下の15 Tenorを使用する。

```math
\mathcal M=
\{1,2,3,4,5,6,7,8,9,10,15,20,25,30,40\}
```

したがって、

```math
n=15
```

である。

### 4.2 Base Tradeの構成

| Type | Base Trade | Hedge Condition |
|---|---|---|
| Outright | 1Y | None |
| Outright | 10Y | None |
| Outright | 25Y | None |
| Curve | 2s3s | PC1-neutral |
| Curve | 3s5s | PC1-neutral |
| Curve | 5s7s | PC1-neutral |
| Curve | 7s10s | PC1-neutral |
| Curve | 10s20s | PC1-neutral |
| Curve | 20s30s | PC1-neutral |
| Curve | 30s40s | PC1-neutral |
| Fly | 2s4s5s | PC1/PC2-neutral |
| Fly | 5s6s7s | PC1/PC2-neutral |
| Fly | 7s8s10s | PC1/PC2-neutral |
| Fly | 7s9s10s | PC1/PC2-neutral |
| Fly | 10s15s20s | PC1/PC2-neutral |

したがって、

```math
\boxed{
T=[T_O\ T_C\ T_F]
}
```

ここで、

```math
T_O\in\mathbb{R}^{15\times3}
```

```math
T_C\in\mathbb{R}^{15\times7}
```

```math
T_F\in\mathbb{R}^{15\times5}
```

である。

### 4.3 フルランク性を保証する構造

OutrightとCurveの10本は、

```math
K=\{1,2,3,5,7,10,20,25,30,40\}
```

の10 Tenorを張る。

7本のCurveは、次のChainを構成している。

```text
2Y -- 3Y -- 5Y -- 7Y -- 10Y -- 20Y -- 30Y -- 40Y
```

このChainと10Y OutrightによってChain上の8次元を張る。

さらに、1Yと25Y Outrightを加えることで、10次元の部分空間を張る。

残る5 Tenorは、

```math
R=\{4,6,8,9,15\}
```

である。

5本のFlyは、それぞれこの5 Tenorのうち1つを固有の中央Legとして持つ。

中央LegのBPVウェイトを1に正規化すると、行・列の並び替えによってTradeMatrixは次のブロック三角構造になる。

```math
\boxed{
T\sim
\begin{pmatrix}
A&H\\
0&I_5
\end{pmatrix}
}
```

ここで `A` は10×10のOutright・Curve基底行列である。

Curveの必要なLegウェイトが非ゼロであり、各Flyの中央Legを非ゼロに正規化できる限り、

```math
\boxed{
\mathrm{rank}(T)=15
}
```

が成立する。

---

## 5. PCAによるBase Tradeのウェイト決定

### 5.1 PCA共分散モデル

15 Tenorの金利変化について、共分散行列を、

```math
\Sigma=\mathrm{Cov}(\Delta y)
```

と定義する。

PCAによって、

```math
\boxed{
\Sigma=V\Lambda V^\top
}
```

と分解する。

固有値は、

```math
\lambda_1\geq\lambda_2\geq\cdots\geq\lambda_n\geq0
```

とする。

JGBイールドカーブでは、主要因子を概念的に、

- PC1：Level
- PC2：Slope
- PC3以降：CurvatureおよびResidual Curve Factors

として解釈する。

ただし、これは経験的な解釈であり、各PCの経済的意味は実際のLoadingによって確認する必要がある。

### 5.2 Outright

OutrightにはPCAヘッジを課さない。

Tenor `a` のOutrightは、

```math
t_O=e_a
```

で定義する。

ここで `e_a` はTenor `a` に対応する単位ベクトルである。

### 5.3 Curve：PC1 Neutrality

例えば5s7s Curveについて、

```math
t_C=e_5+we_7
```

とする。

PC1-neutralityは、

```math
v_1^\top t_C=0
```

である。

これより、

```math
v_{1,5}+wv_{1,7}=0
```

したがって、

```math
\boxed{
w=-\frac{v_{1,5}}{v_{1,7}}
}
```

となる。

このウェイトは、機械的なBPV-neutral ratioではなく、実際のPCA Level Factorに対するHedge Ratioである。

### 5.4 Fly：PC1/PC2 Neutrality

例えば7s9s10s Flyについて、中央Legを1に正規化する。

```math
t_F=w_7e_7+e_9+w_{10}e_{10}
```

PC1とPC2の両方に対して中立となるよう、

```math
\boxed{
\begin{aligned}
v_1^\top t_F&=0\\
v_2^\top t_F&=0
\end{aligned}
}
```

を課す。

したがって、

```math
\begin{pmatrix}
v_{1,7}&v_{1,10}\\
v_{2,7}&v_{2,10}
\end{pmatrix}
\begin{pmatrix}
w_7\\
w_{10}
\end{pmatrix}
=
-
\begin{pmatrix}
v_{1,9}\\
v_{2,9}
\end{pmatrix}
```

となる。

左辺の2×2行列が正則であれば、Wingのウェイトは一意に決まる。

### 5.5 PCA Hedgeによるリスク構造

| Base Trade | PC1 | PC2 | PC3以降 |
|---|---|---|---|
| Outright | Exposureあり得る | Exposureあり得る | Exposureあり得る |
| Curve | 0 | Exposureあり得る | Exposureあり得る |
| Fly | 0 | 0 | Exposureあり得る |

このゼロ構造が、PCA Hedge Ratioを使用する主要なモチベーションである。

**Curveが純粋なPC2ポジション、Flyが純粋なPC3ポジションになるわけではない。**

特にFlyは、PC3以降の複数のリスクファクターを保有し得る。

---

## 6. Tenor・Trade・PCAの座標変換

### 6.1 BPVのTrade Basis変換

```math
\boxed{
x=T^{-1}b
}
```

`x` は各Base Tradeの保有量を表す。

### 6.2 金利変化のTrade Basis変換

Base TradeのBPV加重金利変化を、

```math
\boxed{
\Delta s=T^\top\Delta y
}
```

と定義する。

これはTradeMatrixの各列と金利変化ベクトルの内積である。

この変換は、TradeMatrixを数学的な基底と見なして金利変化の座標を求める操作とは異なる。

ポートフォリオP&Lと整合する**Trade Spread Change**を構築するための変換である。

### 6.3 Trade Covariance

```math
\boxed{
\Sigma_T=T^\top\Sigma T
}
```

これはOutright・Curve・FlyのBPV加重金利変化の共分散行列である。

### 6.4 PCA Exposure

```math
\boxed{
q=V^\top b
}
```

### 6.5 PCA Factor Shock

```math
\boxed{
\Delta f=V^\top\Delta y
}
```

### 6.6 PCA Factor Covariance

```math
\boxed{
Q=V^\top\Sigma V
}
```

`V` が同じ `Σ` を固有値分解した行列であれば、

```math
\boxed{
Q=\Lambda
}
```

となる。

一方、異なる推定期間のPCAと共分散行列を組み合わせる場合、一般には `Q` の非対角成分がゼロにならない。

### 6.7 P&Lの整合性

Tenor Spaceでは、

```math
\Delta PL=-b^\top\Delta y
```

Trade Spaceでは、

```math
\Delta PL=-x^\top\Delta s
```

PCA Spaceでは、

```math
\Delta PL=-q^\top\Delta f
```

したがって、

```math
\boxed{
-b^\top\Delta y
=
-x^\top\Delta s
=
-q^\top\Delta f
}
```

が成立する。

すべて同一のポートフォリオP&Lを異なる空間で表現している。

---

## 7. PCAとTradeMatrixを接続するLoading Matrix

TradeMatrixとPCAを結び付けるため、

```math
\boxed{
L=V^\top T
}
```

と定義する。

`L` の次元は、

```math
L\in\mathbb{R}^{n\times n}
```

であり、

- 行：PC
- 列：Base Trade

である。

各成分は、

```math
\boxed{
L_{ik}=v_i^\top t_k
}
```

である。

すなわち、`L` は**Base Trade 1単位あたりのPCA Factor Loading**を表す。

Portfolio全体のPC Exposureは、

```math
\boxed{
q=Lx
}
```

となる。

PCA Neutralityを利用してTradeMatrixを構築しているため、

```math
L=
\begin{pmatrix}
L_{1,O}&0&0\\
L_{2,O}&L_{2,C}&0\\
L_{R,O}&L_{R,C}&L_{R,F}
\end{pmatrix}
```

という構造が現れる。

ここで `R` はPC3以降を表す。

これは、CurveがPC1 Exposureを保有せず、FlyがPC1・PC2 Exposureを保有しないことを表している。

---

## 8. Signed PCA Exposure Matrix

Base Tradeの保有量を反映したPCA Exposureを定義する。

```math
\boxed{
E_{ki}=x_kL_{ik}
}
```

行列形式では、

```math
\boxed{
E=\mathrm{diag}(x)T^\top V
}
```

となる。

**Eの行はBase Trade、列はPCである。**

| Base Trade | PC1 | PC2 | PC3 | ... |
|---|---:|---:|---:|---:|
| 1Y Outright | Exposure | Exposure | Exposure | ... |
| 10Y Outright | Exposure | Exposure | Exposure | ... |
| 25Y Outright | Exposure | Exposure | Exposure | ... |
| Curve | 0 | Exposure | Exposure | ... |
| Fly | 0 | 0 | Exposure | ... |

この行列は、各Base TradeがどのPCに対してどちら向きのポジションを取っているかを表す。

また、PCごとのExposureをTrade方向に合計すると、Portfolio全体のPC Exposureを再現する。

```math
\boxed{
q_i=\sum_{k=1}^{n}E_{ki}
}
```

行列形式では、

```math
\boxed{
q=E^\top\mathbf1
}
```

となる。

---

## 9. Portfolio VaR

ポートフォリオの分散は、

```math
\boxed{
\sigma_P^2=b^\top\Sigma b
}
```

である。

Trade Spaceに変換すると、

```math
\sigma_P^2=x^\top\Sigma_Tx
```

PCA Spaceでは、

```math
\sigma_P^2=q^\top Qq
```

となる。

したがって、

```math
\boxed{
b^\top\Sigma b
=
x^\top\Sigma_Tx
=
q^\top Qq
}
```

が成立する。

同一の `Σ` を用いたPCAであれば、

```math
\boxed{
\sigma_P^2=
\sum_{i=1}^{n}\lambda_iq_i^2
}
```

となる。

平均損失をゼロとする線形・分散共分散VaRモデルでは、

```math
\boxed{
VaR_\alpha=z_\alpha\sigma_P
}
```

である。

ここで `z` はVaR Multiplierであり、Excel UDFではデフォルトを1とする。

正規分布を仮定する場合の代表的な値は以下の通り。

| Confidence Level | z |
|---|---:|
| 95% | 1.644854 |
| 99% | 2.326348 |

これは金利リスクの一次近似によるParametric VaRであり、Convexity、非線形商品、流動性リスク、極端な市場環境を完全に捉えるものではない。

---

## 10. Euler VaR Decomposition

### 10.1 Tenor別VaR Contribution

VaRを、

```math
VaR=z\sqrt{b^\top\Sigma b}
```

とする。

Tenor `j` のBPVに関する偏微分は、

```math
\frac{\partial VaR}{\partial b_j}
=
\frac{z}{\sigma_P}(\Sigma b)_j
```

である。

Euler Decompositionより、Tenor別Component VaRは、

```math
\boxed{
C_j^{Tenor}
=
\frac{z}{\sigma_P}
b_j(\Sigma b)_j
}
```

となる。

すべてのTenorについて合計すると、

```math
\boxed{
\sum_j C_j^{Tenor}=VaR
}
```

が成立する。

### 10.2 Base Trade別VaR Contribution

Trade SpaceでのPortfolio VaRは、

```math
VaR=z\sqrt{x^\top\Sigma_Tx}
```

である。

Base Trade `k` に関する偏微分は、

```math
\frac{\partial VaR}{\partial x_k}
=
\frac{z}{\sigma_P}(\Sigma_Tx)_k
```

となる。

したがって、Base Trade別Component VaRは、

```math
\boxed{
C_k^{Trade}
=
\frac{z}{\sigma_P}
x_k(\Sigma_Tx)_k
}
```

である。

合計すると、

```math
\boxed{
\sum_kC_k^{Trade}=VaR
}
```

となる。

ここで、Component VaRは各Base Tradeを単独保有した場合のStandalone VaRとは異なる。

Portfolio全体のリスクとの相関・相殺を反映した寄与であり、負の値も取り得る。

---

## 11. PCA × Tradeの二次元VaR分解

### 11.1 定義

PCA Factor別のリスクを、さらにBase Trade別に分解する。

Signed PCA Exposureは、

```math
E_{ki}=x_kL_{ik}
```

である。

PCA座標系におけるFactor Covarianceを、

```math
Q=V^\top\Sigma V
```

と定義し、

```math
\boxed{
m=Qq
}
```

とする。

このとき、PC `i` に対するBase Trade `k` のEuler VaR Contributionは、

```math
\boxed{
A_{ki}
=
\frac{z}{\sigma_P}E_{ki}m_i
}
```

となる。

すなわち、

```math
\boxed{
A_{ki}
=
\frac{z}{\sigma_P}
x_k(V^\top T)_{ik}
(Qq)_i
}
```

である。

同一の `Σ` を固有値分解したPCAでは、

```math
Q=\Lambda
```

なので、

```math
\boxed{
A_{ki}
=
\frac{z}{\sigma_P}
\lambda_iq_ix_k(V^\top T)_{ik}
}
```

となる。

### 11.2 行列形式

**行＝Base Trade、列＝PC** とする。

```math
\boxed{
A=
E\mathrm{diag}
\left(
\frac{z}{\sigma_P}Qq
\right)
}
```

| Base Trade | PC1 | PC2 | PC3 | ... | Trade Total |
|---|---:|---:|---:|---:|---:|
| 1Y Outright | A₁₁ | A₁₂ | A₁₃ | ... | Trade 1 VaR |
| 10Y Outright | A₂₁ | A₂₂ | A₂₃ | ... | Trade 2 VaR |
| 25Y Outright | A₃₁ | A₃₂ | A₃₃ | ... | Trade 3 VaR |
| Curve | 0 | A | A | ... | Curve VaR |
| Fly | 0 | 0 | A | ... | Fly VaR |
| **Factor Total** | **PC1 VaR** | **PC2 VaR** | **PC3 VaR** | ... | **Portfolio VaR** |

### 11.3 Trade方向に合計

各行を合計すると、Base Trade別Component VaRになる。

```math
\boxed{
\sum_i A_{ki}=C_k^{Trade}
}
```

### 11.4 PC方向に合計

各列を合計すると、PC別Component VaRになる。

```math
\boxed{
\sum_k A_{ki}
=
\frac{z}{\sigma_P}q_i(Qq)_i
}
```

同一の共分散行列を対角化するPCAなら、

```math
\boxed{
C_i^{PC}
=
\frac{z}{\sigma_P}\lambda_iq_i^2
}
```

となる。

### 11.5 全体の加法整合性

```math
\boxed{
\sum_k\sum_i A_{ki}=VaR
}
```

この恒等式によって、PCAによるFactor RiskとTrader ModelによるTrade Riskを、同一のリスク寄与行列から分析できる。

---

## 12. VaR Contributionの符号とトレード方向

### 12.1 Component VaRの符号はLong/Shortを示さない

同一の共分散行列を用いたPCAについて、

```math
A_{ki}
=
\frac{z\lambda_i}{\sigma_P}q_iE_{ki}
```

である。

固有値が正の場合、VaR寄与の符号は、

```math
\mathrm{sign}(A_{ki})
=
\mathrm{sign}(q_iE_{ki})
```

によって決まる。

| Portfolio PC Exposure | Trade PC Exposure | VaR Contribution | 解釈 |
|---|---|---|---|
| 正 | 正 | 正 | リスクを増加 |
| 正 | 負 | 負 | リスクを相殺 |
| 負 | 正 | 負 | リスクを相殺 |
| 負 | 負 | 正 | リスクを増加 |

したがって、

**正のComponent VaRはLong Positionを意味しない。**

また、

**負のComponent VaRはShort Positionを意味しない。**

Component VaRはPortfolio Riskへの寄与を表し、ポジション方向そのものを表すわけではない。

実際、ポートフォリオ全体のポジションを反転しても、平均ゼロの分散共分散モデルによるEuler VaR Contributionは不変である。

### 12.2 Trade Directionの定義

トレードの方向性はSigned Exposureによって確認する。

#### Outright

正のBPVをLong Durationとする。

```math
b^{(k)}=x_kt_k
```

該当年限のSigned BPVが正ならLong、負ならShortである。

#### Curve

Curveについては、PC2 Loadingの符号規約を固定した上で、PC2 Exposureによって方向性を把握できる。

ただし、PC2は市場全体のSlope Factorであり、特定の2 Tenor Spreadと完全に同じではない。

より厳密には、対象CurveのSteepening Shockを `h` とし、

```math
\boxed{
\Delta PL_k=-x_kt_k^\top h
}
```

を評価する。

Steepening Shockに対して利益が出るならSteepener、損失が出るならFlattenerである。

#### Fly

Flyについては、高次PCの符号より、TradeMatrixのLeg構造を用いる方が経済的に明確である。

中央Legを+1に正規化し、両Wingが負である場合、

- `x > 0`：Long Belly / Short Wings
- `x < 0`：Short Belly / Long Wings

と解釈できる。

Wingの符号が想定と異なる場合は、実際のSigned BPV構造を確認する必要がある。

---

## 13. Portfolio VaR Shockによる経済的解釈

Component VaRを、実際の金利変動と結び付けて解釈する。

### 13.1 VaR Shock

ポートフォリオの損失は、

```math
Loss=b^\top\Delta y
```

である。

Portfolio VaRに対応する金利変動方向を、

```math
\boxed{
\Delta y^*
=
\frac{z}{\sigma_P}\Sigma b
}
```

と定義する。

これは、正定値の共分散行列に対するMahalanobis距離 `z` の楕円体上で損失を最大化する方向である。

```math
\begin{aligned}
Loss^*
&=b^\top\Delta y^*\\
&=\frac{z}{\sigma_P}b^\top\Sigma b\\
&=z\sigma_P
\end{aligned}
```

したがって、

```math
\boxed{
Loss^*=VaR
}
```

となる。

このShockは将来の金利予測ではなく、共分散モデルと現在のポートフォリオから導かれるリスクシナリオである。

### 13.2 Trade別のVaR Shock損失

Base Trade `k` のBPVは、

```math
b^{(k)}=x_kt_k
```

である。

VaR Shockによる損失は、

```math
Loss_k^*
=
(b^{(k)})^\top\Delta y^*
```

であり、

```math
\boxed{
Loss_k^*=C_k^{Trade}
}
```

が成立する。

したがって、Euler Component VaRは、**Portfolio VaR Shockにおける各Base Tradeの損失**として理解できる。

- 正のComponent VaR：そのShockで損失を出す。
- 負のComponent VaR：そのShockで利益を出し、Portfolio Riskを相殺する。

### 13.3 PCA座標系でのVaR Shock

```math
\Delta f^*=V^\top\Delta y^*
```

とすると、

```math
\boxed{
\Delta f^*
=
\frac{z}{\sigma_P}Qq
}
```

となる。

同一の共分散行列に基づくPCAでは、

```math
\boxed{
\Delta f_i^*
=
\frac{z\lambda_iq_i}{\sigma_P}
}
```

である。

これを使うと、PC×TradeのVaR Contributionは、

```math
\boxed{
A_{ki}=E_{ki}\Delta f_i^*
}
```

と表せる。

これは本フレームワークにおける、最も重要な経済的解釈の一つである。

```text
Signed PCA Exposure
     E(k,i)
        │
        │
        ×
        │
        ▼
Portfolio VaR Factor Shock
     Δf*(i)
        │
        ▼
Factor × Trade VaR Contribution
     A(k,i)
```

**Tradeがどちらの方向にリスクを持っているかと、Portfolioがどちらの金利変動で損失を被るかを掛け合わせることで、VaR Contributionの符号を解釈できる。**

---

## 14. 数値例：2 TenorのVaR分解

計算の整合性を確認するため、簡単な2 Tenorモデルを考える。

```math
b=
\begin{pmatrix}
2\\
1
\end{pmatrix}
```

```math
\Sigma=
\begin{pmatrix}
4&0\\
0&1
\end{pmatrix}
```

```math
T=
\begin{pmatrix}
1&1\\
0&1
\end{pmatrix}
```

```math
V=I
```

とする。

### 14.1 Trade Exposure

```math
x=T^{-1}b
=
\begin{pmatrix}
1\\
1
\end{pmatrix}
```

### 14.2 PCA Exposure

```math
q=V^\top b
=
\begin{pmatrix}
2\\
1
\end{pmatrix}
```

### 14.3 Signed PCA Exposure Matrix

```math
E=\mathrm{diag}(x)T^\top V
```

より、

```math
\boxed{
E=
\begin{pmatrix}
1&0\\
1&1
\end{pmatrix}
}
```

となる。

行がTrade、列がPCである。

### 14.4 Portfolio VaR

```math
\sigma_P^2=b^\top\Sigma b=17
```

したがって、`z=1` の場合、

```math
\boxed{
VaR=\sqrt{17}
}
```

である。

### 14.5 PCA × Trade VaR Matrix

```math
Qq=
\begin{pmatrix}
8\\
1
\end{pmatrix}
```

なので、

```math
\boxed{
A=
\frac{1}{\sqrt{17}}
\begin{pmatrix}
8&0\\
8&1
\end{pmatrix}
}
```

となる。

全要素を合計すると、

```math
\sum_{k,i}A_{ki}
=
\frac{17}{\sqrt{17}}
=
\sqrt{17}
```

となり、Portfolio VaRと一致する。

---

## 15. Excel VBA UDFへの対応

本フレームワークの入力は、基本的に以下の3つである。

```math
\boxed{
b,\quad\Sigma,\quad T
}
```

PCA関連の計算では、共分散行列から求めた固有ベクトル `V` を追加で使用する。`V` は独立した市場入力ではなく、原則として `Σ` から導出されるモデルパラメータである。

### 15.1 UDF一覧

| UDF | 数式 | 出力 |
|---|---|---|
| `TradeVector(T,b)` | `T⁻¹b` | Trade × 1 |
| `TradeCov(T,Cov)` | `TᵀΣT` | Trade × Trade |
| `PortfolioVaR(Cov,b,z)` | `z√(bᵀΣb)` | Scalar |
| `VaRAllocate(Cov,b,z)` | Tenor Euler VaR | Tenor × 1 |
| `TradeVaRAllocate(Cov,b,T,z)` | Trade Euler VaR | Trade × 1 |
| `PCTradeExposure(V,T,b)` | `diag(x)TᵀV` | Trade × PC |
| `PCAVaRAllocate(Cov,b,V,T,z)` | `E diag(zQq/σ)` | Trade × PC |
| `PCAVaRDebug(Cov,b,V,T,z)` | PCA VaR計算診断 | Text |

`z` は省略可能であり、デフォルトは1とする。

### 15.2 Excel標準関数による金利変化の変換

Trade Spread Change：

```excel
=MMULT(TRANSPOSE(T),YieldChange)
```

PCA Factor Change：

```excel
=MMULT(TRANSPOSE(EigenMatrix),YieldChange)
```

Portfolio VaR Shock（`z=1`）：

```excel
=MMULT(Cov,b)/PortfolioVaR(Cov,b)
```

### 15.3 配列の方向

すべての入力・出力において、次の方向を統一する。

| Matrix | Rows | Columns |
|---|---|---|
| `T` | Tenor | Trade |
| `V` | Tenor | PC |
| `L` | PC | Trade |
| `E` | Trade | PC |
| `A` | Trade | PC |
| `Σ` | Tenor | Tenor |
| `ΣT` | Trade | Trade |
| `Q` | PC | PC |

特に、`PCTradeExposure` と `PCAVaRAllocate` はどちらも、**行＝Trade、列＝PC** とする。

---

## 16. リスク計算の検証項目

実際のリスク管理に導入する前に、以下の恒等式と数値的性質を確認する。

### 16.1 Trade Reconstruction

```math
\boxed{
b=Tx
}
```

### 16.2 PCA Reconstruction

```math
\boxed{
q=V^\top b
}
```

```math
\boxed{
b=Vq
}
```

### 16.3 PCA Orthonormality

```math
\boxed{
V^\top V=I
}
```

### 16.4 PCA Covariance

```math
\boxed{
Q=V^\top\Sigma V
}
```

同一共分散行列に基づくPCAなら、

```math
Q\approx\Lambda
```

である。

### 16.5 Variance Reconciliation

```math
\boxed{
b^\top\Sigma b
=
x^\top\Sigma_Tx
=
q^\top Qq
}
```

### 16.6 Signed Exposure Reconciliation

```math
\boxed{
q_i=\sum_kE_{ki}
}
```

### 16.7 VaR Attribution Reconciliation

```math
\boxed{
\begin{aligned}
VaR
&=\sum_jC_j^{Tenor}\\
&=\sum_kC_k^{Trade}\\
&=\sum_k\sum_iA_{ki}
\end{aligned}
}
```

### 16.8 PCA Hedge Validation

Curveについて、

```math
v_1^\top t_C\approx0
```

Flyについて、

```math
v_1^\top t_F\approx0
```

```math
v_2^\top t_F\approx0
```

を確認する。

また、TradeMatrixのRank・条件数も継続的に確認する。

---

## 17. 実務上の留意点

### 17.1 PCA推定期間への依存

PCAの固有ベクトルは推定期間に依存する。

特に高次PCは不安定になりやすく、特定のPC3やPC4の経済的解釈を固定することは危険である。

そのため、PC1/PC2の中立性に加えて、PC3以降のResidual Subspace全体の安定性を評価することが望ましい。

### 17.2 PCA Window変更によるTradeMatrixの変化

PCA Hedge Ratioを更新すると、TradeMatrixも変化する。

```math
T=T(V)
```

したがって、

```math
x=T(V)^{-1}b
```

も変化する。

同一のBPVポートフォリオでも、PCA推定期間を変更することで、Trade ExposureとTrade別Component VaRの数値が変わる可能性がある。

これは実際のポジション変化ではなく、リスク帰属に用いる基底の変更によって生じる場合がある。

### 17.3 共分散行列とPCAの整合性

異なる推定期間の `Σ` と `V` を組み合わせる場合、

```math
Q=V^\top\Sigma V
```

は必ずしも対角行列にならない。

この場合、PC別VaR分解には、

```math
\boxed{
A_{ki}
=
\frac{z}{\sigma_P}E_{ki}(Qq)_i
}
```

を使用し、非対角成分を含む完全なFactor Covarianceを考慮する。

その場合、Factorは分析対象の共分散行列に対して必ずしも無相関ではない。

### 17.4 固有ベクトルの符号規約

PCAの固有ベクトルには符号の任意性がある。

そのため、PC Exposureの正負を経済的に解釈する場合は、符号規約を固定する必要がある。

例えば、

- PC1：概ね金利上昇方向を正
- PC2：長短金利差の拡大方向を正

と定義できる。

ただし、固有ベクトルの符号を反転しても、対応するExposureとFactor Shockを同時に反転させれば、VaR Contributionそのものは不変である。

### 17.5 Trade Basisへの依存

Portfolio VaRは、正則なTradeMatrixによる座標変換に対して不変である。

しかし、Trade別Component VaRはTradeMatrixの選択に依存する。

したがって、Trade別VaR Attributionは唯一の客観的な分解ではなく、選択したTrader Modelに基づくリスク帰属である。

PCAヘッジによるTradeMatrixの設計は、VaRの数値自体を改善するためではなく、そのリスク帰属を経済的に解釈しやすくするためのものである。

### 17.6 追加的なリスク

本フレームワークは一次の金利リスクを対象としている。

実務では必要に応じて、

- Convexity Risk
- Futures Basis Risk
- Repo / Funding Risk
- Liquidity Risk
- Stress Scenario Risk
- Model Risk

などを別途管理する必要がある。

また、共分散行列が半正定値であること、金利変化とBPVの単位が整合していること、PCAの対象TenorとBPVの対象Tenorが一致していることが前提となる。

---

## 18. 最終的なリスク管理フロー

```text
                  MARKET DATA
                      │
                      ▼
        Yield Changes / Covariance
                      │
                      ▼
              Σ = Cov(Δy)
                      │
                      ▼
              PCA Decomposition
               Σ = VΛV^T
                      │
                      ▼
           PCA Hedge Ratio Calculation
                      │
                      ▼
          TradeMatrix Construction
                     T
                      │
                      │
                Portfolio BPV
                     b
                      │
             ┌────────┴────────┐
             ▼                 ▼
       Trade Exposure      PCA Exposure
         x = T^(-1)b         q = V^Tb
             │                 │
             ▼                 ▼
       PCA Trade Loading   Factor Covariance
          L = V^TT          Q = V^TΣV
             │                 │
             ▼                 ▼
        Signed PCA       Factor Marginal
         Exposure              Risk
       E = diag(x)L^T         m = Qq
             │                 │
             └────────┬────────┘
                      ▼
          Factor × Trade VaR Matrix
                A(k,i)
                      │
             ┌────────┴────────┐
             ▼                 ▼
        Trade VaR            PC VaR
         Allocation         Allocation
             │                 │
             └────────┬────────┘
                      ▼
                Portfolio VaR
```

## 19. 主要な数式のまとめ

### Trade Basis

```math
\boxed{
\begin{aligned}
b&=Tx\\
x&=T^{-1}b\\
\Delta s&=T^\top\Delta y\\
\Sigma_T&=T^\top\Sigma T
\end{aligned}
}
```

### PCA Basis

```math
\boxed{
\begin{aligned}
\Sigma&=V\Lambda V^\top\\
q&=V^\top b\\
\Delta f&=V^\top\Delta y\\
Q&=V^\top\Sigma V
\end{aligned}
}
```

### PCA × Trade Exposure

```math
\boxed{
\begin{aligned}
L&=V^\top T\\
E&=\mathrm{diag}(x)T^\top V
\end{aligned}
}
```

### Portfolio Risk

```math
\boxed{
\begin{aligned}
\sigma_P^2
&=b^\top\Sigma b\\
&=x^\top\Sigma_Tx\\
&=q^\top Qq
\end{aligned}
}
```

```math
\boxed{
VaR=z\sigma_P
}
```

### Euler VaR Attribution

```math
\boxed{
\begin{aligned}
C_j^{Tenor}
&=\frac{z}{\sigma_P}b_j(\Sigma b)_j\\[6pt]
C_k^{Trade}
&=\frac{z}{\sigma_P}x_k(\Sigma_Tx)_k\\[6pt]
A_{ki}
&=\frac{z}{\sigma_P}E_{ki}(Qq)_i
\end{aligned}
}
```

### Portfolio VaR Shock

```math
\boxed{
\begin{aligned}
\Delta y^*&=\frac{z}{\sigma_P}\Sigma b\\
\Delta f^*&=\frac{z}{\sigma_P}Qq\\
A_{ki}&=E_{ki}\Delta f_i^*
\end{aligned}
}
```

### 最終的な加法整合性

```math
\boxed{
\sum_jC_j^{Tenor}
=
\sum_kC_k^{Trade}
=
\sum_k\sum_iA_{ki}
=
VaR
}
```

---

## 20. 結論

本フレームワークの目的は、PCA・TradeMatrix・VaRを独立したリスク管理ツールとして使用するのではなく、同一の線形代数的構造の中で統合することである。

その中心となる考え方は次の3点である。

**第一に、TradeMatrixはポートフォリオのBPVを、トレーダーが直感的に理解できるBase Tradeへ変換するための基底である。**

**第二に、PCA Hedge RatioによってBase Tradeを定義することで、各Tradeが保有する統計的リスクファクターを階層的に制御できる。**

**第三に、Signed PCA ExposureとPortfolio VaR Shockを組み合わせることで、各Tradeの方向性とEuler VaR Contributionを一貫して解釈できる。**

最終的には、

```math
\boxed{
\underbrace{E_{ki}}_{\text{Signed PCA Exposure}}
\times
\underbrace{\Delta f_i^*}_{\text{Portfolio VaR Shock}}
=
\underbrace{A_{ki}}_{\text{Factor × Trade VaR}}
}
```

という関係が成立する。

これにより、単に「どのTradeがどれだけVaRを持っているか」だけでなく、**「どのTradeが、どのPCに対して、どちら向きのリスクを取り、それがPortfolio全体のVaRにどのように寄与しているか」**を体系的に理解できる。

これが、PCA Risk ModelとTrader Modelを統合したJGB相対価値トレーディングのリスク管理フレームワークである。
