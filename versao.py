"""Versão do Conciliador — fonte única pra tela e pro pacote distribuído.

O formato é a data da build (``AAAA.MM.DD``), igual ao nome do ZIP gerado
pelo ``publicar.bat`` (``Conciliador_AAAA-MM-DD.zip``). Assim o suporte
cruza o que o operador lê na tela com o pacote que foi publicado: se a
máquina mostra ``v2026.09.14`` e a pasta de rede tem a ``2026.09.15``,
aquela máquina não rodou a atualização.

O ``publicar.bat`` reescreve a linha do VERSAO com a data do dia antes de
chamar o PyInstaller — não é preciso editar na mão.
"""

VERSAO = "2026.09.15"


def rotulo() -> str:
    """Texto curto exibido no título da janela e na barra do topo."""
    return f"v{VERSAO}"
