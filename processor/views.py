from django.shortcuts import render
from django.core.files.storage import FileSystemStorage
import pandas as pd
import os
from .utils_interim import generate_interim_report
from .utils_word import generate_final_word_report
# Fix the Import Error by using the correct function name
from .utils_compare import compare_pre_post_scans 

def home(request):
    """Operations Dashboard"""
    return render(request, 'processor/home.html')

def interim_pre(request):
    """Phase 1: Generate Initial Report"""
    context = {}
    if request.method == 'POST' and request.FILES.getlist('scans'):
        fs = FileSystemStorage()
        files = [fs.save(f.name, f) for f in request.FILES.getlist('scans')]
        action = request.POST.get('action')
        
        excel_name, csv_name = generate_interim_report(files, prefix="Pre_Remediation")
        
        if action == 'generate_excel':
            context['excel_url'] = fs.url(excel_name)
        elif action == 'generate_csv':
            context['csv_url'] = fs.url(csv_name)
            
        return render(request, 'processor/upload_pre.html', context)
    return render(request, 'processor/upload_pre.html')

def interim_post(request):
    """Phase 2: Final Consolidated Report View"""
    context = {}
    if request.method == 'POST' and request.FILES.get('pre_file'):
        fs = FileSystemStorage()
        
        # 1. Save uploaded Round 1 report and Round 2 scans
        pre_file = fs.save(request.FILES['pre_file'].name, request.FILES['pre_file'])
        post_files = [fs.save(f.name, f) for f in request.FILES.getlist('post_scans')]
        
        # 2. Use Phase 1 script to process post-scans into a clean baseline
        _, post_csv_temp = generate_interim_report(post_files, prefix="Post_Baseline")
        
        # 3. Run Comparison (returns both names)
        excel_out_req = "Consolidated_VA_Report.xlsx"
        excel_name, csv_name = compare_pre_post_scans(
            fs.path(pre_file), 
            fs.path(post_csv_temp), 
            excel_out_req
        )
        
        # 4. Pass URLs to template context
        context['excel_url'] = fs.url(excel_name)
        context['csv_url'] = fs.url(csv_name)
        
    return render(request, 'processor/upload_post.html', context)

def word_report(request):
    """Phase 3: Export to Word"""
    if request.method == 'POST' and request.FILES.get('report_file'):
        fs = FileSystemStorage()
        up_file = request.FILES.get('report_file')
        path = fs.path(fs.save(up_file.name, up_file))
        
        if up_file.name.endswith('.xlsx'):
            # It will now look for the final tab if coming from Phase 2
            try:
                df = pd.read_excel(path, sheet_name='InterimReport_VA_final')
            except:
                df = pd.read_excel(path, sheet_name='InterimReport_VA')
        else:
            df = pd.read_csv(path)
            
        word_doc = generate_final_word_report(df)
        return render(request, 'processor/upload_word.html', {'word_url': fs.url(word_doc)})
    return render(request, 'processor/upload_word.html')


# Go to the very bottom of views.py

def test_reporting_dashboard(request):
    return render(request, 'processor/reporting_dashboard.html')  # <--- MUST BE INDENTED