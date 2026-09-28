# Change-PCA Residual Subspace Framework for JGB Relative Value

## 1. Yield Curve and Change PCA

Yield curve:

\[
\mathbf y_t
=
\begin{pmatrix}
y_{1,t}\\
y_{2,t}\\
\vdots\\
y_{n,t}
\end{pmatrix}
\in\mathbb R^n
\]

Daily change:

\[
\Delta\mathbf y_t
=
\mathbf y_t-\mathbf y_{t-1}
\]

Change covariance:

\[
\Sigma_\Delta
=
\operatorname{Cov}(\Delta\mathbf y_t)
\]

PCA:

\[
\boxed{
\Sigma_\Delta
=
V\Lambda V^\mathsf T
}
\]

where

\[
V
=
\begin{pmatrix}
\mathbf v_1&
\mathbf v_2&
\cdots&
\mathbf v_n
\end{pmatrix},
\qquad
V^\mathsf TV=I
\]

and

\[
\Lambda
=
\operatorname{diag}
(\lambda_1,\lambda_2,\ldots,\lambda_n),
\qquad
\lambda_1\ge\lambda_2\ge\cdots\ge\lambda_n.
\]

PC scores:

\[
\boxed{
\mathbf f_t
=
V^\mathsf T\Delta\mathbf y_t
}
\]

\[
f_{i,t}
=
\mathbf v_i^\mathsf T\Delta\mathbf y_t
\]

Reconstruction:

\[
\Delta\mathbf y_t
=
V\mathbf f_t
=
\sum_{i=1}^{n}
\mathbf v_i f_{i,t}
\]

---

## 2. Systematic Subspace

Assume:

\[
PC1 \approx \text{Level}
\]

\[
PC2 \approx \text{Slope}
\]

Define

\[
V_2
=
\begin{pmatrix}
\mathbf v_1&
\mathbf v_2
\end{pmatrix}.
\]

Projection onto PC1-PC2:

\[
\boxed{
P_2
=
V_2V_2^\mathsf T
}
\]

Residual projector:

\[
\boxed{
Q_2
=
I-P_2
=
I-V_2V_2^\mathsf T
}
\]

Therefore

\[
\boxed{
\mathbf x
=
P_2\mathbf x
+
Q_2\mathbf x
}
\]

with

\[
P_2\mathbf x
=
\text{systematic component}
\]

and

\[
Q_2\mathbf x
=
\text{residual component}.
\]

Since

\[
I
=
\sum_{i=1}^{n}
\mathbf v_i\mathbf v_i^\mathsf T,
\]

\[
\boxed{
Q_2
=
\sum_{i=3}^{n}
\mathbf v_i\mathbf v_i^\mathsf T
}
\]

and

\[
\boxed{
Q_2\mathbf x
=
\sum_{i=3}^{n}
\mathbf v_i
(\mathbf v_i^\mathsf T\mathbf x)
}
\]

Hence the framework does **not** require identification of a specific PC3, PC4, ..., PCn trade.

---

## 3. Residual Calculation

### 3.1 Daily Residual / Innovation

Systematic daily move:

\[
\widehat{\Delta\mathbf y}_t
=
P_2\Delta\mathbf y_t
\]

Residual daily move:

\[
\boxed{
\mathbf u_t
=
\Delta\mathbf y_t
-
\widehat{\Delta\mathbf y}_t
}
\]

Therefore

\[
\boxed{
\mathbf u_t
=
Q_2\Delta\mathbf y_t
}
\]

or equivalently,

\[
\boxed{
\mathbf u_t
=
\left(
I-V_2V_2^\mathsf T
\right)
\Delta\mathbf y_t
}
\]

---

### 3.2 Historical Residual State

For yield levels:

\[
\boxed{
\mathbf R_t
=
Q_2(\mathbf y_t-\mathbf c)
}
\]

where \(\mathbf c\) is an optional historical reference curve.

If \(\mathbf c=0\),

\[
\boxed{
\mathbf R_t
=
Q_2\mathbf y_t
}
\]

Under a fixed PCA basis:

\[
\begin{aligned}
\Delta\mathbf R_t
&=
\mathbf R_t-\mathbf R_{t-1}\\
&=
Q_2(\mathbf y_t-\mathbf y_{t-1})\\
&=
Q_2\Delta\mathbf y_t
\end{aligned}
\]

Therefore

\[
\boxed{
\Delta\mathbf R_t
=
\mathbf u_t
}
\]

Interpretation:

\[
\boxed{
\mathbf R_t=\text{Residual State}
}
\]

\[
\boxed{
\mathbf u_t=\text{Residual Innovation}
}
\]

---

## 4. PCA Hedge

Signed BPV vector:

\[
\mathbf b
=
\begin{pmatrix}
b_1\\
b_2\\
\vdots\\
b_n
\end{pmatrix}
\]

Approximate P\&L:

\[
\boxed{
\Delta P_t
\approx
-\mathbf b^\mathsf T
\Delta\mathbf y_t
}
\]

Using PCA:

\[
\Delta\mathbf y_t
=
V\mathbf f_t
\]

so

\[
\Delta P_t
=
-\mathbf b^\mathsf TV\mathbf f_t.
\]

Define PCA exposure:

\[
\boxed{
\mathbf g
=
V^\mathsf T\mathbf b
}
\]

Then

\[
\boxed{
\Delta P_t
=
-\mathbf g^\mathsf T\mathbf f_t
}
\]

with

\[
\boxed{
g_i
=
\mathbf v_i^\mathsf T\mathbf b
}
\]

---

## 5. PC1-PC2 Neutral Hedge

PC1-PC2 neutrality:

\[
\boxed{
V_2^\mathsf T\mathbf b
=
\mathbf 0
}
\]

i.e.

\[
\boxed{
\mathbf v_1^\mathsf T\mathbf b=0
}
\]

\[
\boxed{
\mathbf v_2^\mathsf T\mathbf b=0
}
\]

Example: 5s7s10s

\[
\mathbf b
=
\begin{pmatrix}
0\\
\vdots\\
b_5\\
\vdots\\
b_7\\
\vdots\\
b_{10}\\
\vdots\\
0
\end{pmatrix}
\]

Solve

\[
V_2^\mathsf T\mathbf b=0
\]

plus one normalization condition, e.g.

\[
b_7=1.
\]

For a 3-leg trade:

\[
3\text{ weights}
-
1\text{ scale}
-
2\text{ hedge constraints}
=
0
\]

remaining degrees of freedom.

Therefore:

\[
\boxed{
\text{3-leg + PC1/PC2 hedge}
\Rightarrow
\text{unique ratio up to scale}
}
\]

---

## 6. Remaining PCA Exposure

After hedge:

\[
\mathbf g
=
V^\mathsf T\mathbf b
\]

with

\[
g_1=g_2=0.
\]

In general:

\[
\boxed{
\mathbf g
=
\begin{pmatrix}
0\\
0\\
g_3\\
g_4\\
\vdots\\
g_n
\end{pmatrix}
}
\]

Hence

\[
\boxed{
PC1/PC2\text{-neutral}
\neq
PC3\text{-pure}
}
\]

and

\[
\boxed{
\mathbf b
=
\sum_{i=3}^{n}
g_i\mathbf v_i
}
\]

---

## 7. Variance-Adjusted PCA Risk

Raw PCA exposure:

\[
g_i
=
\mathbf v_i^\mathsf T\mathbf b
\]

is not sufficient for risk comparison.

Because

\[
\operatorname{Var}(\mathbf f_t)
=
\Lambda,
\]

factor \(i\) has variance

\[
\operatorname{Var}(f_i)
=
\lambda_i.
\]

Portfolio variance:

\[
\begin{aligned}
\operatorname{Var}(\Delta P)
&=
\mathbf b^\mathsf T
\Sigma_\Delta
\mathbf b\\
&=
\mathbf b^\mathsf T
V\Lambda V^\mathsf T
\mathbf b\\
&=
(V^\mathsf T\mathbf b)^\mathsf T
\Lambda
(V^\mathsf T\mathbf b)\\
&=
\mathbf g^\mathsf T
\Lambda
\mathbf g
\end{aligned}
\]

Therefore

\[
\boxed{
\operatorname{Var}(\Delta P)
=
\sum_{i=1}^{n}
\lambda_i g_i^2
}
\]

Factor variance contribution:

\[
\boxed{
RC_i
=
\lambda_i g_i^2
}
\]

Factor risk contribution share:

\[
\boxed{
w_i^{risk}
=
\frac{\lambda_i g_i^2}
{\sum_{j=1}^{n}\lambda_j g_j^2}
}
\]

and

\[
\sum_{i=1}^{n}w_i^{risk}=1.
\]

For a PC1-PC2-neutral position:

\[
\boxed{
\operatorname{Var}(\Delta P)
=
\sum_{i=3}^{n}
\lambda_i g_i^2
}
\]

Higher-PC aggregate risk:

\[
\boxed{
R_{\text{Residual}}
=
\sum_{i=3}^{n}
\lambda_i g_i^2
}
\]

or, after separating PC3:

\[
\boxed{
R_{\text{Tail}}
=
\sum_{i=4}^{n}
\lambda_i g_i^2
}
\]

PC3 risk share:

\[
\boxed{
\rho_3
=
\frac{\lambda_3g_3^2}
{\sum_{i=3}^{n}\lambda_i g_i^2}
}
\]

Tail risk share:

\[
\boxed{
\rho_{\text{Tail}}
=
\frac{
\sum_{i=4}^{n}\lambda_i g_i^2
}{
\sum_{i=3}^{n}\lambda_i g_i^2
}
}
\]

with

\[
\rho_3+\rho_{\text{Tail}}=1.
\]

---

## 8. Higher-PC Instability

Individual higher PCs may rotate when eigenvalues are close:

\[
\lambda_i
\approx
\lambda_{i+1}.
\]

However the residual space is

\[
\boxed{
\mathcal R
=
\operatorname{span}
(\mathbf v_1,\mathbf v_2)^\perp
}
\]

and

\[
\boxed{
Q_2
=
I-V_2V_2^\mathsf T
}
\]

depends only on the top-2 subspace.

If

\[
\widetilde V_2
=
V_2R
\]

for orthogonal \(R\),

\[
RR^\mathsf T=I,
\]

then

\[
\widetilde V_2
\widetilde V_2^\mathsf T
=
V_2V_2^\mathsf T.
\]

Therefore

\[
\boxed{
\widetilde P_2=P_2
}
\]

and

\[
\boxed{
\widetilde Q_2=Q_2
}
\]

even if PC1 and PC2 rotate internally.

The relevant stability object is therefore

\[
\boxed{
\operatorname{span}(PC1,PC2)
}
\]

rather than individual higher PCs.

---

## 9. Multiple PCA Windows

For estimation window \(w\):

\[
\Sigma_\Delta^{(w)}
=
\operatorname{Cov}^{(w)}
(\Delta\mathbf y)
\]

\[
\Sigma_\Delta^{(w)}
=
V^{(w)}
\Lambda^{(w)}
V^{(w)\mathsf T}
\]

and

\[
P_2^{(w)}
=
V_2^{(w)}
V_2^{(w)\mathsf T}.
\]

Example:

\[
w
\in
\{63,126,252,504\}.
\]

Compare

\[
\boxed{
P_2^{63},
P_2^{126},
P_2^{252},
P_2^{504}
}
\]

rather than only individual eigenvectors.

For hedge vector \(\mathbf b^{(63)}\),

\[
V_2^{(63)\mathsf T}
\mathbf b^{(63)}
=
0
\]

but generally

\[
\boxed{
V_2^{(252)\mathsf T}
\mathbf b^{(63)}
\neq0
}
\]

which provides a measure of hedge-model uncertainty.

---

## 10. RV Workflow

### A. Structure-First

\[
\text{Trade Idea}
\rightarrow
\text{Structure}
\rightarrow
\boxed{
V_2^\mathsf T\mathbf b=0
}
\rightarrow
\text{PCA Hedge}
\]

Example:

\[
5s7s10s
\rightarrow
\text{PC1/PC2-neutral weights}
\]

---

### B. Daily Distortion Monitor

\[
\boxed{
\mathbf u_t
=
Q_2\Delta\mathbf y_t
}
\]

Detect unusual daily curve movements.

---

### C. Historical Residual State

\[
\boxed{
\mathbf R_t
=
Q_2(\mathbf y_t-\mathbf c)
}
\]

Monitor historical residual curve shape.

With fixed basis:

\[
\boxed{
\Delta\mathbf R_t
=
\mathbf u_t
}
\]

---

## 11. Trader Decision Layer

PCA outputs:

\[
\boxed{
\mathbf R_t,\quad
\mathbf u_t,\quad
\mathbf b,\quad
\mathbf g,\quad
\lambda_i g_i^2
}
\]

Trade decision:

\[
E_t[\text{Return}]
=
F
\left(
\mathbf R_t,
\mathbf u_t,
\text{Flow}_t,
\text{Supply/Demand}_t,
\text{Macro}_t,
\text{Policy}_t,
\text{Liquidity}_t
\right)
\]

PCA does not determine:

\[
\boxed{
\text{Mean Reversion}
}
\]

or

\[
\boxed{
\text{Trend Continuation}
}
\]

The trader determines the direction.

---

# 12. Framework Summary

\[
\boxed{
\Delta\mathbf y
\xrightarrow{\operatorname{Cov}}
\Sigma_\Delta
\xrightarrow{\operatorname{PCA}}
V,\Lambda
}
\]

\[
\boxed{
V_2
\rightarrow
P_2=V_2V_2^\mathsf T
\rightarrow
Q_2=I-P_2
}
\]

\[
\boxed{
\mathbf u_t
=
Q_2\Delta\mathbf y_t
}
\]

\[
\boxed{
\mathbf R_t
=
Q_2(\mathbf y_t-\mathbf c)
}
\]

\[
\boxed{
\mathbf u_t
=
\Delta\mathbf R_t
}
\]

Trade structure:

\[
\boxed{
V_2^\mathsf T\mathbf b=0
}
\]

PCA exposure:

\[
\boxed{
\mathbf g=V^\mathsf T\mathbf b
}
\]

Variance-adjusted risk:

\[
\boxed{
RC_i
=
\lambda_i g_i^2
}
\]

Total risk:

\[
\boxed{
\operatorname{Var}(\Delta P)
=
\sum_i\lambda_i g_i^2
}
\]

---

# 13. Core Principle

\[
\boxed{
\text{PCA does not provide the RV answer.}
}
\]

\[
\boxed{
\text{PCA provides the coordinate system for RV analysis.}
}
\]

More specifically:

\[
\boxed{
\text{Change PCA}
\rightarrow
\text{Systematic Subspace}
\rightarrow
\text{Residual Subspace}
}
\]

\[
\boxed{
\text{Residual State / Innovation}
\rightarrow
\text{Trader Interpretation}
\rightarrow
\text{Trade Structure}
\rightarrow
\text{PCA Hedge}
}
\]
