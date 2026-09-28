# Change-PCA Residual Subspace Framework for JGB Relative Value

## 1. Objective

PCA is used for:

- curve-state representation
- systematic / residual decomposition
- PCA hedge-ratio construction
- factor-risk decomposition

PCA is **not** used to determine whether a distortion will mean-revert or continue trending.

Final trade direction is determined separately using:

- flow
- supply / demand
- positioning
- liquidity
- BOJ / policy
- macro regime
- trader discretion

---

# 2. Yield Curve

Yield curve:

```math
\mathbf{y}_t
=
\begin{pmatrix}
y_{1,t} \\
y_{2,t} \\
\vdots \\
y_{n,t}
\end{pmatrix}
\in
\mathbb{R}^n
```

Daily yield change:

```math
\Delta \mathbf{y}_t
=
\mathbf{y}_t
-
\mathbf{y}_{t-1}
```

---

# 3. Change PCA

Estimate the covariance matrix from yield changes:

```math
\Sigma_{\Delta}
=
\mathrm{Cov}
\left(
\Delta \mathbf{y}_t
\right)
```

Eigenvalue decomposition:

```math
\Sigma_{\Delta}
=
V
\Lambda
V^{\mathrm T}
```

where

```math
V
=
\begin{pmatrix}
\mathbf{v}_1 &
\mathbf{v}_2 &
\cdots &
\mathbf{v}_n
\end{pmatrix}
```

and

```math
V^{\mathrm T}V
=
I
```

Eigenvalue matrix:

```math
\Lambda
=
\mathrm{diag}
\left(
\lambda_1,
\lambda_2,
\ldots,
\lambda_n
\right)
```

with

```math
\lambda_1
\ge
\lambda_2
\ge
\cdots
\ge
\lambda_n
```

PC scores:

```math
\mathbf{f}_t
=
V^{\mathrm T}
\Delta\mathbf{y}_t
```

For PC \(i\):

```math
f_{i,t}
=
\mathbf{v}_i^{\mathrm T}
\Delta\mathbf{y}_t
```

Reconstruction:

```math
\Delta\mathbf{y}_t
=
V\mathbf{f}_t
```

or

```math
\Delta\mathbf{y}_t
=
\sum_{i=1}^{n}
\mathbf{v}_i f_{i,t}
```

---

# 4. Systematic Subspace

Empirically for the JGB curve:

```math
PC1
\approx
\mathrm{Level}
```

```math
PC2
\approx
\mathrm{Slope}
```

Define:

```math
V_2
=
\begin{pmatrix}
\mathbf{v}_1 &
\mathbf{v}_2
\end{pmatrix}
```

Projection onto the PC1-PC2 subspace:

```math
P_2
=
V_2
V_2^{\mathrm T}
```

Residual projector:

```math
Q_2
=
I-P_2
```

Therefore:

```math
Q_2
=
I
-
V_2
V_2^{\mathrm T}
```

For any vector \(\mathbf{x}\):

```math
\mathbf{x}
=
P_2\mathbf{x}
+
Q_2\mathbf{x}
```

where:

```math
P_2\mathbf{x}
=
\mathrm{Systematic\ Component}
```

and

```math
Q_2\mathbf{x}
=
\mathrm{Residual\ Component}
```

---

# 5. Residual Subspace

Because:

```math
I
=
\sum_{i=1}^{n}
\mathbf{v}_i
\mathbf{v}_i^{\mathrm T}
```

we have:

```math
Q_2
=
\sum_{i=3}^{n}
\mathbf{v}_i
\mathbf{v}_i^{\mathrm T}
```

Therefore:

```math
Q_2\mathbf{x}
=
\sum_{i=3}^{n}
\mathbf{v}_i
\left(
\mathbf{v}_i^{\mathrm T}
\mathbf{x}
\right)
```

Residual subspace:

```math
\mathcal{R}
=
\mathrm{span}
\left(
\mathbf{v}_1,
\mathbf{v}_2
\right)^{\perp}
```

The framework therefore does not require a trade to be identified with a specific PC3, PC4, ..., PCn.

---

# 6. Daily Residual Innovation

Systematic daily yield move:

```math
\widehat{\Delta\mathbf{y}}_t
=
P_2
\Delta\mathbf{y}_t
```

Residual:

```math
\mathbf{u}_t
=
\Delta\mathbf{y}_t
-
\widehat{\Delta\mathbf{y}}_t
```

Therefore:

```math
\mathbf{u}_t
=
Q_2
\Delta\mathbf{y}_t
```

or:

```math
\mathbf{u}_t
=
\left(
I
-
V_2V_2^{\mathrm T}
\right)
\Delta\mathbf{y}_t
```

Interpretation:

```math
\mathbf{u}_t
=
\mathrm{Residual\ Innovation}
```

This measures the part of today's curve movement that cannot be explained by PC1-PC2.

---

# 7. Historical Residual State

Apply the Change-PCA systematic subspace to yield levels:

```math
\mathbf{R}_t
=
Q_2
\left(
\mathbf{y}_t
-
\mathbf{c}
\right)
```

where \(\mathbf{c}\) is an optional reference curve.

If no reference curve is used:

```math
\mathbf{R}_t
=
Q_2
\mathbf{y}_t
```

Thus:

```math
\mathbf{R}_t
=
\left(
I
-
V_2V_2^{\mathrm T}
\right)
\mathbf{y}_t
```

Interpretation:

```math
\mathbf{R}_t
=
\mathrm{Residual\ State}
```

This measures the current yield-curve shape after removing the PC1-PC2 component.

---

# 8. Relationship between State and Innovation

For a fixed PCA basis:

```math
\mathbf{R}_t
=
Q_2\mathbf{y}_t
```

Then:

```math
\mathbf{R}_t
-
\mathbf{R}_{t-1}
=
Q_2
\left(
\mathbf{y}_t
-
\mathbf{y}_{t-1}
\right)
```

Therefore:

```math
\Delta\mathbf{R}_t
=
Q_2
\Delta\mathbf{y}_t
```

Hence:

```math
\Delta\mathbf{R}_t
=
\mathbf{u}_t
```

So:

```math
\mathbf{R}_t
=
\mathrm{Residual\ State}
```

```math
\mathbf{u}_t
=
\mathrm{Residual\ Innovation}
```

Methodologically:

```math
\mathrm{Innovation}
=
\Delta
\left(
\mathrm{State}
\right)
```

---

# 9. Historical Normalization

Historical mean:

```math
\overline{\mathbf{R}}
=
\frac{1}{T}
\sum_{t=1}^{T}
\mathbf{R}_t
```

For tenor \(j\):

```math
Z_{j,t}
=
\frac{
R_{j,t}
-
\overline{R}_j
}{
\sigma_j
}
```

This is interpreted as historical extremeness, not automatically as fair-value mispricing.

---

# 10. PCA Hedge

Let the signed BPV vector be:

```math
\mathbf{b}
=
\begin{pmatrix}
b_1 \\
b_2 \\
\vdots \\
b_n
\end{pmatrix}
```

First-order P&L:

```math
\Delta P_t
\approx
-
\mathbf{b}^{\mathrm T}
\Delta\mathbf{y}_t
```

Since:

```math
\Delta\mathbf{y}_t
=
V\mathbf{f}_t
```

we have:

```math
\Delta P_t
=
-
\mathbf{b}^{\mathrm T}
V
\mathbf{f}_t
```

Define PCA exposure:

```math
\mathbf{g}
=
V^{\mathrm T}
\mathbf{b}
```

Therefore:

```math
\Delta P_t
=
-
\mathbf{g}^{\mathrm T}
\mathbf{f}_t
```

For each PC:

```math
g_i
=
\mathbf{v}_i^{\mathrm T}
\mathbf{b}
```

---

# 11. PC1-PC2 Neutral Hedge

PC1-PC2 neutrality:

```math
V_2^{\mathrm T}
\mathbf{b}
=
\mathbf{0}
```

Equivalent to:

```math
\mathbf{v}_1^{\mathrm T}
\mathbf{b}
=
0
```

and:

```math
\mathbf{v}_2^{\mathrm T}
\mathbf{b}
=
0
```

---

# 12. Example: 5s7s10s Fly

Full BPV vector:

```math
\mathbf{b}
=
\begin{pmatrix}
0 \\
\vdots \\
b_5 \\
\vdots \\
b_7 \\
\vdots \\
b_{10} \\
\vdots \\
0
\end{pmatrix}
```

Hedge constraints:

```math
\mathbf{v}_1^{\mathrm T}
\mathbf{b}
=
0
```

```math
\mathbf{v}_2^{\mathrm T}
\mathbf{b}
=
0
```

Normalization example:

```math
b_7
=
1
```

For a three-leg trade:

```math
3
-
1
-
2
=
0
```

where:

- 3 = leg weights
- 1 = overall scale
- 2 = PC hedge constraints

Therefore the hedge ratio is unique up to overall scale.

---

# 13. Remaining PCA Exposure

After PCA hedge:

```math
\mathbf{g}
=
V^{\mathrm T}
\mathbf{b}
```

with:

```math
g_1
=
0
```

```math
g_2
=
0
```

but generally:

```math
\mathbf{g}
=
\begin{pmatrix}
0 \\
0 \\
g_3 \\
g_4 \\
\vdots \\
g_n
\end{pmatrix}
```

Therefore:

```math
\mathrm{PC1/PC2\ Neutral}
\neq
\mathrm{PC3\ Pure}
```

and:

```math
\mathbf{b}
=
\sum_{i=3}^{n}
g_i
\mathbf{v}_i
```

The position belongs to the residual subspace.

---

# 14. Variance-Adjusted PCA Risk

Raw PC exposure:

```math
g_i
=
\mathbf{v}_i^{\mathrm T}
\mathbf{b}
```

does not by itself represent factor risk.

Since:

```math
\mathrm{Var}
\left(
\mathbf{f}_t
\right)
=
\Lambda
```

we have:

```math
\mathrm{Var}
\left(
f_i
\right)
=
\lambda_i
```

Portfolio variance:

```math
\mathrm{Var}
\left(
\Delta P
\right)
=
\mathbf{b}^{\mathrm T}
\Sigma_{\Delta}
\mathbf{b}
```

Substituting the PCA decomposition:

```math
\mathrm{Var}
\left(
\Delta P
\right)
=
\mathbf{b}^{\mathrm T}
V
\Lambda
V^{\mathrm T}
\mathbf{b}
```

Using:

```math
\mathbf{g}
=
V^{\mathrm T}
\mathbf{b}
```

we obtain:

```math
\mathrm{Var}
\left(
\Delta P
\right)
=
\mathbf{g}^{\mathrm T}
\Lambda
\mathbf{g}
```

Therefore:

```math
\mathrm{Var}
\left(
\Delta P
\right)
=
\sum_{i=1}^{n}
\lambda_i
g_i^2
```

---

# 15. Factor Variance Contribution

Factor \(i\) variance contribution:

```math
RC_i
=
\lambda_i
g_i^2
```

Total variance:

```math
RC_{\mathrm{Total}}
=
\sum_{i=1}^{n}
RC_i
```

or:

```math
RC_{\mathrm{Total}}
=
\sum_{i=1}^{n}
\lambda_i
g_i^2
```

Risk contribution share:

```math
w_i^{\mathrm{risk}}
=
\frac{
\lambda_i
g_i^2
}{
\sum_{j=1}^{n}
\lambda_j
g_j^2
}
```

and:

```math
\sum_{i=1}^{n}
w_i^{\mathrm{risk}}
=
1
```

For a PC1-PC2-neutral trade:

```math
\mathrm{Var}
\left(
\Delta P
\right)
=
\sum_{i=3}^{n}
\lambda_i
g_i^2
```

---

# 16. Residual Risk

Residual-subspace risk:

```math
R_{\mathrm{Residual}}
=
\sum_{i=3}^{n}
\lambda_i
g_i^2
```

If PC3 is separated:

```math
R_{\mathrm{Tail}}
=
\sum_{i=4}^{n}
\lambda_i
g_i^2
```

PC3 risk share:

```math
\rho_3
=
\frac{
\lambda_3
g_3^2
}{
\sum_{i=3}^{n}
\lambda_i
g_i^2
}
```

Tail risk share:

```math
\rho_{\mathrm{Tail}}
=
\frac{
\sum_{i=4}^{n}
\lambda_i
g_i^2
}{
\sum_{i=3}^{n}
\lambda_i
g_i^2
}
```

Therefore:

```math
\rho_3
+
\rho_{\mathrm{Tail}}
=
1
```

---

# 17. Higher-PC Instability

If:

```math
\lambda_i
\approx
\lambda_{i+1}
```

individual eigenvectors may rotate significantly across estimation samples.

However, suppose:

```math
\widetilde{V}_2
=
V_2
R
```

where \(R\) is orthogonal:

```math
RR^{\mathrm T}
=
I
```

Then:

```math
\widetilde{P}_2
=
\widetilde{V}_2
\widetilde{V}_2^{\mathrm T}
```

and:

```math
\widetilde{P}_2
=
V_2
RR^{\mathrm T}
V_2^{\mathrm T}
```

Therefore:

```math
\widetilde{P}_2
=
V_2
V_2^{\mathrm T}
```

Hence:

```math
\widetilde{P}_2
=
P_2
```

and:

```math
\widetilde{Q}_2
=
Q_2
```

Therefore the relevant stability object is the top-two subspace:

```math
\mathrm{span}
\left(
PC1,
PC2
\right)
```

rather than individual higher PCs.

---

# 18. PC2-PC3 Boundary

A key diagnostic is the eigenvalue gap:

```math
\lambda_2
-
\lambda_3
```

If:

```math
\lambda_2
\gg
\lambda_3
```

the top-two subspace tends to be more clearly separated.

If:

```math
\lambda_2
\approx
\lambda_3
```

the distinction between systematic and residual subspaces may become unstable.

---

# 19. Multiple Estimation Windows

For window \(w\):

```math
\Sigma_{\Delta}^{(w)}
=
\mathrm{Cov}^{(w)}
\left(
\Delta\mathbf{y}
\right)
```

PCA:

```math
\Sigma_{\Delta}^{(w)}
=
V^{(w)}
\Lambda^{(w)}
V^{(w)\mathrm T}
```

Systematic projection:

```math
P_2^{(w)}
=
V_2^{(w)}
V_2^{(w)\mathrm T}
```

Residual projection:

```math
Q_2^{(w)}
=
I
-
P_2^{(w)}
```

Example windows:

```math
w
\in
\{
63,
126,
252,
504
\}
```

Compare:

```math
P_2^{(63)},
\quad
P_2^{(126)},
\quad
P_2^{(252)},
\quad
P_2^{(504)}
```

rather than only comparing individual eigenvectors.

---

# 20. Cross-Model Hedge Robustness

Suppose a hedge vector is constructed using the 63-day PCA:

```math
V_2^{(63)\mathrm T}
\mathbf{b}^{(63)}
=
0
```

It does not necessarily satisfy:

```math
V_2^{(252)\mathrm T}
\mathbf{b}^{(63)}
=
0
```

Cross-model residual exposure:

```math
\mathbf{h}^{(63 \rightarrow 252)}
=
V_2^{(252)\mathrm T}
\mathbf{b}^{(63)}
```

This provides a direct measure of PCA hedge model uncertainty.

---

# 21. Practical RV Workflow

## A. Structure-First

Trader identifies a structure:

```math
5s7s10s
```

Then solve:

```math
V_2^{\mathrm T}
\mathbf{b}
=
0
```

Result:

```math
\mathrm{Trade\ Idea}
\rightarrow
\mathrm{PCA\ Hedge}
```

PCA is used only for weight construction.

---

## B. Daily Distortion Monitor

Calculate:

```math
\mathbf{u}_t
=
Q_2
\Delta\mathbf{y}_t
```

Use this to identify unusual relative moves.

---

## C. Historical Residual Monitor

Calculate:

```math
\mathbf{R}_t
=
Q_2
\left(
\mathbf{y}_t-\mathbf{c}
\right)
```

Use this to identify historically extreme residual curve shapes.

---

## D. Joint State / Innovation Monitor

Monitor jointly:

```math
\left(
\mathbf{R}_t,
\mathbf{u}_t
\right)
```

with:

```math
\mathbf{u}_t
=
\Delta\mathbf{R}_t
```

---

## E. Trade Selection

Trader chooses a structure based on:

- residual state
- residual innovation
- flow
- supply / demand
- macro
- policy
- liquidity
- positioning

---

## F. Hedge Construction

For the selected structure:

```math
V_2^{\mathrm T}
\mathbf{b}
=
0
```

---

## G. Remaining-Risk Analysis

Calculate:

```math
\mathbf{g}
=
V^{\mathrm T}
\mathbf{b}
```

and:

```math
RC_i
=
\lambda_i
g_i^2
```

---

# 22. Framework Architecture

```math
\Delta\mathbf{y}
\rightarrow
\Sigma_{\Delta}
\rightarrow
V,\Lambda
```

```math
V_2
\rightarrow
P_2
=
V_2V_2^{\mathrm T}
```

```math
P_2
\rightarrow
Q_2
=
I-P_2
```

Daily residual:

```math
\mathbf{u}_t
=
Q_2
\Delta\mathbf{y}_t
```

Historical residual:

```math
\mathbf{R}_t
=
Q_2
\left(
\mathbf{y}_t-\mathbf{c}
\right)
```

State / innovation relation:

```math
\Delta\mathbf{R}_t
=
\mathbf{u}_t
```

PCA hedge:

```math
V_2^{\mathrm T}
\mathbf{b}
=
0
```

PCA exposure:

```math
\mathbf{g}
=
V^{\mathrm T}
\mathbf{b}
```

Variance-adjusted PC risk:

```math
RC_i
=
\lambda_i
g_i^2
```

Total risk:

```math
\mathrm{Var}
\left(
\Delta P
\right)
=
\sum_{i=1}^{n}
\lambda_i
g_i^2
```

---

# 23. Role Separation

PCA:

```math
\mathrm{PCA}
=
\mathrm{State\ Representation}
+
\mathrm{Risk\ Decomposition}
+
\mathrm{Hedge\ Construction}
```

Trader:

```math
\mathrm{Trader}
=
\mathrm{Economic\ Interpretation}
+
\mathrm{Trade\ Direction}
```

PCA does not determine:

```math
\mathrm{Mean\ Reversion}
```

or:

```math
\mathrm{Trend\ Continuation}
```

---

# 24. Core Principle

```math
\mathrm{Change\ PCA}
\rightarrow
\mathrm{Systematic\ Subspace}
\rightarrow
\mathrm{Residual\ Subspace}
```

```math
\mathrm{Residual\ State}
+
\mathrm{Residual\ Innovation}
\rightarrow
\mathrm{Trader\ Interpretation}
```

```math
\mathrm{Trader\ Interpretation}
\rightarrow
\mathrm{Trade\ Structure}
\rightarrow
\mathrm{PCA\ Hedge}
```

Final principle:

```math
\mathrm{PCA}
\neq
\mathrm{RV\ Answer}
```

```math
\mathrm{PCA}
=
\mathrm{Coordinate\ System\ for\ RV\ Analysis}
```
