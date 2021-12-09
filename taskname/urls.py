from django.contrib import admin
from django.urls import path, include
from todolist_app import views as todolist_views



urlpatterns = [
    path('admin/', admin.site.urls),
    path('', todolist_views.index1a, name='index1a'),
    #path('', todolist_views.indexapi, name='indexapi'),

    path('mul/', todolist_views.FileFieldView, name='FileFieldView'),
    path('gstr2a', todolist_views.index2A, name='index2A')

    ]
