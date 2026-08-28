from django.db import models
from django.contrib.auth.models import User

class Unidade(models.Model):
    nome = models.CharField(max_length=150, unique=True)
    codigo = models.CharField(max_length=50, blank=True, null=True)
    ativo = models.BooleanField(default=True)
    criado_em = models.DateTimeField(auto_now_add=True)
    class Meta:
        verbose_name = "Unidade"
        verbose_name_plural = "Unidades"
        ordering = ["nome"]
    def __str__(self):
        return self.nome

class PerfilUsuario(models.Model):
    PAPEL_CHOICES = (
        ("GESTOR", "Gestor Master (Acesso Total)"),
        ("CONTROLLER", "Controller (Validação Técnica)"),
        ("OPERADOR", "Operador de Unidade (Restrito às Unidades Vinculadas)"),
    )
    user = models.OneToOneField(User, on_delete=models.CASCADE, related_name="perfil")
    papel = models.CharField(max_length=20, choices=PAPEL_CHOICES, default="GESTOR")
    polos_ids = models.TextField(blank=True, default="", help_text="IDs separados por vírgula")

    class Meta:
        verbose_name = "Perfil de Acesso (Cargo)"
        verbose_name_plural = "Perfis de Acesso (Cargos)"

    def __str__(self):
        return f"{self.user.username} - {self.get_papel_display()}"

    def get_unidades_ids(self):
        if not self.polos_ids:
            return []
        try:
            return [int(x.strip()) for x in str(self.polos_ids).split(",") if x.strip().isdigit()]
        except Exception:
            return []

class Categoria(models.Model):
    TIPO_CHOICES = (
        ("OUT", "Money OUT (Despesa)"),
        ("IN", "Money IN (Receita)"),
    )
    nome = models.CharField(max_length=100)
    tipo = models.CharField(max_length=5, choices=TIPO_CHOICES, default="OUT")
    unidades = models.ManyToManyField(Unidade, blank=True, related_name="categorias", verbose_name="Unidades com Acesso")
    disponivel_todas = models.BooleanField(default=True, verbose_name="Disponível para todas as unidades")

    class Meta:
        verbose_name = "Categoria"
        verbose_name_plural = "Categorias"
        ordering = ["nome"]

    def __str__(self):
        return f"[{self.tipo}] {self.nome}"

class ContaBancaria(models.Model):
    nome = models.CharField(max_length=120, help_text="Ex: Itaú PJ Matriz, Bradesco Piracicaba, Caixa Físico")
    banco = models.CharField(max_length=100, blank=True, null=True, help_text="Ex: Banco Itaú, Bradesco")
    agencia = models.CharField(max_length=30, blank=True, null=True)
    conta = models.CharField(max_length=50, blank=True, null=True)
    titular = models.CharField(max_length=150, blank=True, null=True)
    chave_pix = models.CharField(max_length=255, blank=True, null=True, verbose_name="Chave Pix da Conta")
    
    permite_gestores = models.BooleanField(default=True, verbose_name="Visível para Gestores Master")
    permite_controller = models.BooleanField(default=True, verbose_name="Visível para Controller")
    permite_operador = models.BooleanField(default=False, verbose_name="Visível para Operador")
    somente_criador = models.BooleanField(default=False, verbose_name="Somente quem criou o registro (Dono)")

    criado_por = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name="contas_criadas")
    ativo = models.BooleanField(default=True)
    criado_em = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Conta Bancária / Fonte"
        verbose_name_plural = "Contas Bancárias / Fontes"
        ordering = ["nome"]

    def __str__(self):
        return self.nome

class ContaPagar(models.Model):
    STATUS_CHOICES = (
        ("PENDENTE", "Pendente"),
        ("APROVADO", "Aprovado"),
        ("PAGO", "Pago"),
        ("REJEITADO", "Rejeitado"),
    )
    unidade = models.ForeignKey(Unidade, on_delete=models.CASCADE, related_name="despesas")
    categoria = models.ForeignKey(Categoria, on_delete=models.SET_NULL, null=True, blank=True, related_name="despesas")
    conta_origem = models.ForeignKey(ContaBancaria, on_delete=models.SET_NULL, null=True, blank=True, related_name="despesas_pagas", verbose_name="Conta de Origem")
    fornecedor = models.CharField(max_length=200)
    descricao = models.CharField(max_length=255, blank=True, null=True)
    dados_pix = models.CharField(max_length=255, blank=True, null=True, verbose_name="Dados Pix")
    valor = models.DecimalField(max_digits=12, decimal_places=2)
    data_vencimento = models.DateField()
    status = models.CharField(max_length=25, choices=STATUS_CHOICES, default="PENDENTE")
    boleto = models.FileField(upload_to="boletos/%Y/%m/", null=True, blank=True)
    aprovacao = models.FileField(upload_to="aprovacoes/%Y/%m/", null=True, blank=True)
    comprovante = models.FileField(upload_to="comprovantes/%Y/%m/", null=True, blank=True)
    criado_por = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name="despesas_criadas")
    
    recorrente = models.BooleanField(default=False)
    parcela_atual = models.IntegerField(default=1)
    total_parcelas = models.IntegerField(default=1, help_text="0 indica custo fixo mensal contínuo")
    grupo_recorrencia_id = models.CharField(max_length=50, blank=True, null=True)

    criado_em = models.DateTimeField(auto_now_add=True)
    atualizado_em = models.DateTimeField(auto_now=True)
    class Meta:
        verbose_name = "Conta a Pagar (Money OUT)"
        verbose_name_plural = "Contas a Pagar (Money OUT)"
        ordering = ["data_vencimento"]
    def __str__(self):
        return f"{self.unidade.nome} - {self.fornecedor} (R$ {self.valor})"

class ContaReceber(models.Model):
    STATUS_IN_CHOICES = (
        ("PREVISTO", "Previsto"),
        ("CONFIRMADO", "Confirmado / Recebido"),
    )
    unidade = models.ForeignKey(Unidade, on_delete=models.CASCADE, related_name="receitas")
    categoria = models.ForeignKey(Categoria, on_delete=models.SET_NULL, null=True, blank=True, related_name="receitas")
    conta_destino = models.ForeignKey(ContaBancaria, on_delete=models.SET_NULL, null=True, blank=True, related_name="receitas_recebidas", verbose_name="Conta de Destino / Depósito")
    origem = models.CharField(max_length=200, default="Repasse Mensalidade")
    descricao = models.CharField(max_length=255, blank=True, null=True)
    valor = models.DecimalField(max_digits=12, decimal_places=2)
    data_previsao = models.DateField()
    status = models.CharField(max_length=25, choices=STATUS_IN_CHOICES, default="PREVISTO")
    
    recorrente = models.BooleanField(default=False)
    parcela_atual = models.IntegerField(default=1)
    total_parcelas = models.IntegerField(default=1, help_text="0 indica receita fixa mensal contínua")
    grupo_recorrencia_id = models.CharField(max_length=50, blank=True, null=True)

    criado_em = models.DateTimeField(auto_now_add=True)
    class Meta:
        verbose_name = "Receita / Entrada (Money IN)"
        verbose_name_plural = "Receitas / Entradas (Money IN)"
        ordering = ["data_previsao"]
    def __str__(self):
        return f"[IN] {self.unidade.nome} - {self.origem} (R$ {self.valor})"