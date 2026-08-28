from django.contrib import admin
from django import forms
from .models import Unidade, PerfilUsuario, Categoria, ContaPagar, ContaReceber, ContaBancaria

class PerfilUsuarioForm(forms.ModelForm):
    unidades_selecionadas = forms.ModelMultipleChoiceField(
        queryset=Unidade.objects.filter(ativo=True),
        widget=forms.CheckboxSelectMultiple,
        required=False,
        label="Unidades Permitidas"
    )

    class Meta:
        model = PerfilUsuario
        fields = ['user', 'papel']

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        if self.instance and self.instance.pk:
            try:
                ids = self.instance.get_unidades_ids()
                if ids:
                    self.fields['unidades_selecionadas'].initial = Unidade.objects.filter(id__in=ids)
            except Exception:
                pass

    def save(self, commit=True):
        perfil = super().save(commit=False)
        selecionadas = self.cleaned_data.get('unidades_selecionadas')
        if selecionadas:
            perfil.polos_ids = ",".join([str(u.id) for u in selecionadas])
        else:
            perfil.polos_ids = ""
        if commit:
            perfil.save()
        return perfil

@admin.register(PerfilUsuario)
class PerfilUsuarioAdmin(admin.ModelAdmin):
    form = PerfilUsuarioForm
    list_display = ("user", "papel", "listar_unidades")
    list_filter = ("papel",)
    search_fields = ("user__username", "user__email")

    def listar_unidades(self, obj):
        try:
            ids = obj.get_unidades_ids()
            if not ids:
                return "Todas (Global / Master)"
            nomes = Unidade.objects.filter(id__in=ids).values_list("nome", flat=True)
            return ", ".join(nomes)
        except Exception:
            return "Erro ao carregar unidades"
    listar_unidades.short_description = "Unidades Vinculadas"

@admin.register(Unidade)
class UnidadeAdmin(admin.ModelAdmin):
    list_display = ("id", "nome", "codigo", "ativo", "criado_em")
    search_fields = ("nome", "codigo")

@admin.register(Categoria)
class CategoriaAdmin(admin.ModelAdmin):
    list_display = ("nome", "tipo", "disponivel_todas")
    list_filter = ("tipo", "disponivel_todas")
    search_fields = ("nome",)
    filter_horizontal = ("unidades",)

@admin.register(ContaBancaria)
class ContaBancariaAdmin(admin.ModelAdmin):
    list_display = ("nome", "banco", "agencia", "conta", "chave_pix", "permite_gestores", "permite_controller", "permite_operador", "somente_criador", "ativo")
    list_filter = ("permite_gestores", "permite_controller", "permite_operador", "somente_criador", "ativo")
    search_fields = ("nome", "banco", "titular", "chave_pix")

@admin.register(ContaPagar)
class ContaPagarAdmin(admin.ModelAdmin):
    list_display = ("unidade", "fornecedor", "valor", "data_vencimento", "status", "conta_origem", "recorrente")
    list_filter = ("status", "unidade", "data_vencimento", "conta_origem")
    search_fields = ("fornecedor", "descricao", "dados_pix")
    date_hierarchy = "data_vencimento"

@admin.register(ContaReceber)
class ContaReceberAdmin(admin.ModelAdmin):
    list_display = ("unidade", "origem", "valor", "data_previsao", "status", "conta_destino", "recorrente")
    list_filter = ("status", "unidade", "data_previsao", "conta_destino")
    search_fields = ("origem", "descricao")