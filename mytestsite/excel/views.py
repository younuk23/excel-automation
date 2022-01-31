from django.http import HttpResponse
from . import excel


def index(request):
    result = excel.bootstrap(request.FILES['file'].file)
    response = HttpResponse(content=result, content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename=Inform.xlsx'
    return response


def ping(request):
    return HttpResponse("pong")
