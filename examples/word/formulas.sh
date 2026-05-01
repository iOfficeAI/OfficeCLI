#!/bin/bash
# Generate complex math/chemistry/physics formula test document
# Usage: ./gen_formulas.sh [officecli path]

CLI="${1:-officecli}"
OUT="$(dirname "$0")/complex_formulas.docx"

rm -f "$OUT"
$CLI create "$OUT"
$CLI open "$OUT"

# ==================== Title ====================
$CLI add "$OUT" /body --type paragraph --prop text="Complex Math/Chemistry/Physics Formula Collection" --prop style=Heading1 --prop align=center

# ==================== I. Algebra ====================
$CLI add "$OUT" /body --type paragraph --prop text="I. Algebra" --prop style=Heading2

$CLI add "$OUT" /body --type paragraph --prop text="1. Quadratic Formula:"
$CLI add "$OUT" /body --type equation --prop 'formula=x = \frac{-b \pm \sqrt{b^{2} - 4ac}}{2a}'

$CLI add "$OUT" /body --type paragraph --prop text="2. Binomial Theorem:"
$CLI add "$OUT" /body --type equation --prop 'formula=(a+b)^{n} = \sum_{k=0}^{n} \binom{n}{k} a^{n-k} b^{k}'

$CLI add "$OUT" /body --type paragraph --prop text="3. Euler's Identity:"
$CLI add "$OUT" /body --type equation --prop 'formula=e^{i\pi} + 1 = 0'

# ==================== II. Calculus ====================
$CLI add "$OUT" /body --type paragraph --prop text="II. Calculus" --prop style=Heading2

$CLI add "$OUT" /body --type paragraph --prop text="4. Limit Definition of Derivative:"
$CLI add "$OUT" /body --type equation --prop 'formula=f^{\prime}(x) = \lim_{\Delta x \rightarrow 0} \frac{f(x + \Delta x) - f(x)}{\Delta x}'

$CLI add "$OUT" /body --type paragraph --prop text="5. Gaussian Integral:"
$CLI add "$OUT" /body --type equation --prop 'formula=\int_{-\infty}^{+\infty} e^{-x^{2}} dx = \sqrt{\pi}'

$CLI add "$OUT" /body --type paragraph --prop text="6. Taylor Series Expansion:"
$CLI add "$OUT" /body --type equation --prop 'formula=f(x) = \sum_{n=0}^{\infty} \frac{f^{(n)}(a)}{n!} (x-a)^{n}'

$CLI add "$OUT" /body --type paragraph --prop text="7. Newton-Leibniz Formula:"
$CLI add "$OUT" /body --type equation --prop 'formula=\int_{a}^{b} f(x) dx = F(b) - F(a)'

$CLI add "$OUT" /body --type paragraph --prop text="8. Triple Integral (Spherical Coordinates):"
$CLI add "$OUT" /body --type equation --prop 'formula=\iiint_{V} f(r, \theta, \phi) r^{2} \sin\theta \, dr \, d\theta \, d\phi'

$CLI add "$OUT" /body --type paragraph --prop text="9. Fourier Transform:"
$CLI add "$OUT" /body --type equation --prop 'formula=\hat{f}(\xi) = \int_{-\infty}^{+\infty} f(x) e^{-2\pi i x \xi} dx'

# ==================== III. Linear Algebra ====================
$CLI add "$OUT" /body --type paragraph --prop text="III. Linear Algebra" --prop style=Heading2

$CLI add "$OUT" /body --type paragraph --prop text="10. Matrix Characteristic Equation:"
$CLI add "$OUT" /body --type equation --prop 'formula=\det(A - \lambda I) = 0'

# ==================== IV. Probability and Statistics ====================
$CLI add "$OUT" /body --type paragraph --prop text="IV. Probability and Statistics" --prop style=Heading2

$CLI add "$OUT" /body --type paragraph --prop text="11. Bayes' Theorem:"
$CLI add "$OUT" /body --type equation --prop 'formula=P(A|B) = \frac{P(B|A) \cdot P(A)}{P(B)}'

$CLI add "$OUT" /body --type paragraph --prop text="12. Normal Distribution PDF:"
$CLI add "$OUT" /body --type equation --prop 'formula=f(x) = \frac{1}{\sigma \sqrt{2\pi}} e^{-\frac{(x-\mu)^{2}}{2\sigma^{2}}}'

$CLI add "$OUT" /body --type paragraph --prop text="13. Variance Formula:"
$CLI add "$OUT" /body --type equation --prop 'formula=\sigma^{2} = \frac{1}{N} \sum_{i=1}^{N} (x_{i} - \mu)^{2}'

# ==================== V. Number Theory and Series ====================
$CLI add "$OUT" /body --type paragraph --prop text="V. Number Theory and Series" --prop style=Heading2

$CLI add "$OUT" /body --type paragraph --prop text="14. Riemann Zeta Function:"
$CLI add "$OUT" /body --type equation --prop 'formula=\zeta(s) = \sum_{n=1}^{\infty} \frac{1}{n^{s}}'

$CLI add "$OUT" /body --type paragraph --prop text="15. Stirling's Approximation:"
$CLI add "$OUT" /body --type equation --prop 'formula=n! \approx \sqrt{2\pi n} \left(\frac{n}{e}\right)^{n}'

# ==================== VI. Chemistry ====================
$CLI add "$OUT" /body --type paragraph --prop text="VI. Chemistry" --prop style=Heading2

$CLI add "$OUT" /body --type paragraph --prop text="16. Copper Sulfate Crystal Dissolution:"
$CLI add "$OUT" /body --type equation --prop 'formula=CuSO_{4} \cdot 5H_{2}O \rightarrow Cu^{2+} + SO_{4}^{2-} + 5H_{2}O'

$CLI add "$OUT" /body --type paragraph --prop text="17. Thermochemical Equation (Methane Combustion):"
$CLI add "$OUT" /body --type equation --prop 'formula=CH_{4}(g) + 2O_{2}(g) \rightarrow CO_{2}(g) + 2H_{2}O(l) \quad \Delta H = -890.3 \, kJ/mol'

$CLI add "$OUT" /body --type paragraph --prop text="18. Chemical Equilibrium Constant Expression:"
$CLI add "$OUT" /body --type equation --prop 'formula=K_{eq} = \frac{[C]^{c} [D]^{d}}{[A]^{a} [B]^{b}}'

$CLI add "$OUT" /body --type paragraph --prop text="19. Esterification Reaction (Reversible):"
$CLI add "$OUT" /body --type equation --prop 'formula=CH_{3}COOH + C_{2}H_{5}OH \rightleftharpoons CH_{3}COOC_{2}H_{5} + H_{2}O'

$CLI add "$OUT" /body --type paragraph --prop text="20. Henderson-Hasselbalch Equation:"
$CLI add "$OUT" /body --type equation --prop 'formula=pH = pK_{a} + \log \frac{[A^{-}]}{[HA]}'

$CLI add "$OUT" /body --type paragraph --prop text="21. Van der Waals Equation:"
$CLI add "$OUT" /body --type equation --prop 'formula=\left(P + \frac{a n^{2}}{V^{2}}\right)(V - nb) = nRT'

$CLI add "$OUT" /body --type paragraph --prop text="22. Arrhenius Equation:"
$CLI add "$OUT" /body --type equation --prop 'formula=k = A e^{-\frac{E_{a}}{RT}}'

# ==================== VII. Physics ====================
$CLI add "$OUT" /body --type paragraph --prop text="VII. Physics" --prop style=Heading2

$CLI add "$OUT" /body --type paragraph --prop text="23. Maxwell's Equations (Differential Form):"
$CLI add "$OUT" /body --type equation --prop 'formula=\nabla \cdot E = \frac{\rho}{\epsilon_{0}}'
$CLI add "$OUT" /body --type equation --prop 'formula=\nabla \cdot B = 0'
$CLI add "$OUT" /body --type equation --prop 'formula=\nabla \times E = -\frac{\partial B}{\partial t}'
$CLI add "$OUT" /body --type equation --prop 'formula=\nabla \times B = \mu_{0} J + \mu_{0} \epsilon_{0} \frac{\partial E}{\partial t}'

$CLI add "$OUT" /body --type paragraph --prop text="24. Einstein Field Equations:"
$CLI add "$OUT" /body --type equation --prop 'formula=R_{\mu\nu} - \frac{1}{2} R g_{\mu\nu} + \Lambda g_{\mu\nu} = \frac{8\pi G}{c^{4}} T_{\mu\nu}'

$CLI add "$OUT" /body --type paragraph --prop text="25. Schrodinger Equation:"
$CLI add "$OUT" /body --type equation --prop 'formula=i\hbar \frac{\partial}{\partial t} \Psi(r, t) = \hat{H} \Psi(r, t)'

$CLI add "$OUT" /body --type paragraph --prop text="26. Dirac Equation:"
$CLI add "$OUT" /body --type equation --prop 'formula=(i\gamma^{\mu} \partial_{\mu} - m) \psi = 0'

$CLI add "$OUT" /body --type paragraph --prop text="27. Euler-Lagrange Equation:"
$CLI add "$OUT" /body --type equation --prop 'formula=\frac{d}{dt} \frac{\partial L}{\partial \dot{q}_{i}} - \frac{\partial L}{\partial q_{i}} = 0'

$CLI add "$OUT" /body --type paragraph --prop text="28. Heisenberg Uncertainty Principle:"
$CLI add "$OUT" /body --type equation --prop 'formula=\Delta x \cdot \Delta p \geq \frac{\hbar}{2}'

$CLI add "$OUT" /body --type paragraph --prop text="29. Planck's Black-Body Radiation Formula:"
$CLI add "$OUT" /body --type equation --prop 'formula=B(\nu, T) = \frac{2h\nu^{3}}{c^{2}} \cdot \frac{1}{e^{\frac{h\nu}{k_{B} T}} - 1}'

$CLI add "$OUT" /body --type paragraph --prop text="30. Lorentz Transformation:"
$CLI add "$OUT" /body --type equation --prop 'formula=t^{\prime} = \gamma \left(t - \frac{vx}{c^{2}}\right), \quad \gamma = \frac{1}{\sqrt{1 - \frac{v^{2}}{c^{2}}}}'

$CLI close "$OUT"

echo "Generated: $OUT"
