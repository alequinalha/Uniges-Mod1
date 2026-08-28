from django.urls import path
from . import views

urlpatterns = [
    path("", views.dashboard_gestor, name="dashboard_gestor"),
    path("tendencia/", views.tendencia_view, name="tendencia"),
    path("login/", views.login_view, name="login"),
    path("logout/", views.logout_view, name="logout"),
    path("alterar-senha/", views.alterar_senha_view, name="alterar_senha"),
    path("extrato/", views.extrato_conta_view, name="extrato_conta"),
    path("extrato/exportar-excel/", views.exportar_extrato_excel, name="exportar_extrato_excel"),
    path("categorias/", views.gerenciar_categorias_view, name="gerenciar_categorias"),
    path("categorias/editar/<int:pk>/", views.editar_categoria_view, name="editar_categoria"),
    path("categorias/deletar/<int:pk>/", views.deletar_categoria_view, name="deletar_categoria"),
    path("contas/", views.gerenciar_contas_view, name="gerenciar_contas"),
    path("contas/editar/<int:pk>/", views.editar_conta_bancaria_view, name="editar_conta_bancaria"),
    path("contas/deletar/<int:pk>/", views.deletar_conta_bancaria_view, name="deletar_conta_bancaria"),
    path("aprovar-lote/", views.aprovar_em_lote, name="aprovar_em_lote"),
    path("criar-despesa/", views.criar_despesa_ajax, name="criar_despesa_ajax"),
    path("editar-despesa/<int:pk>/", views.editar_despesa_ajax, name="editar_despesa_ajax"),
    path("criar-receita/", views.criar_receita_ajax, name="criar_receita_ajax"),
    path("editar-receita/<int:pk>/", views.editar_receita_ajax, name="editar_receita_ajax"),
    path("deletar-despesa/<int:pk>/", views.deletar_despesa, name="deletar_despesa"),
    path("deletar-receita/<int:pk>/", views.deletar_receita, name="deletar_receita"),
    path("exportar-excel/", views.exportar_excel, name="exportar_excel"),
    path("exportar-mes-json/", views.exportar_contas_mes, name="exportar_contas_mes"),
    path("importar-mes-json/", views.importar_contas_mes, name="importar_contas_mes"),
]