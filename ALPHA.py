# No terminal:  pip install numpy matplotlib

import numpy as np
import matplotlib.pyplot as plt

def get_weight_fraction(n_range, alpha):
    """Calcula a soma das frações mássicas para um intervalo de carbonos."""
    wn_sum = 0
    for n in n_range:
        wn_sum += n * (1 - alpha)**2 * alpha**(n-1)
    return wn_sum

# Vetor de probabilidade alpha (de 0 a 1)
alpha = np.linspace(0.001, 0.999, 200)

# Definição das faixas do gráfico
curves = {
    r'$C_1$': [1],
    r'$C_2 - C_4$': range(2, 5),
    r'$C_5 - C_{11}$': range(5, 12),
    r'$C_{12} - C_{20}$': range(12, 21),
    r'$C_{20} - C_{30}$': range(20, 31),
    r'$C_{30}^{+}$': None # Tratado separadamente como o resíduo
}

plt.figure(figsize=(8, 7))

# Plotagem das faixas definidas
for label, n_range in curves.items():
    if label == r'$C_{30}^{+}$':
        # C31+ é 1 menos a soma de C1 até C30
        y = 1 - get_weight_fraction(range(1, 31), alpha)
        plt.plot(alpha, y, label=label, linewidth=3, color='olive')
    else:
        y = get_weight_fraction(n_range, alpha)
        plt.plot(alpha, y, label=label, linewidth=2.5)

# Formatação padrão dissertação UPE
plt.xlabel(r'Probabilidade de crescimento da cadeia ($\alpha$)', fontsize=12)
plt.ylabel(r'Fração mássica (wt%)', fontsize=12)
#plt.title('Distribuição de Produtos Anderson-Schulz-Flory (ASF)')
plt.grid(True, linestyle='--', alpha=0.6)
plt.legend(frameon=True, loc='upper right', bbox_to_anchor=(0.8, 0.8))

# Salva em alta resolução na pasta do seu projeto
plt.savefig('distribuicao_ASF_final.png', dpi=300, bbox_inches='tight')
plt.show()