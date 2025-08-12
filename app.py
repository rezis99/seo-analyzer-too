from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import requests
import xml.etree.ElementTree as ET
from urllib.parse import urlparse
import time
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
import io
from datetime import datetime

app = Flask(__name__)
CORS(app)

# Simple configuration
MAX_WORKERS = 5
REQUEST_TIMEOUT = 10

def create_session():
    """Create HTTP session with headers"""
    session = requests.Session()
    session.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    })
    return session

def extract_seo_data(soup, url):
    """Extract basic SEO elements from webpage"""
    try:
        # Title
        title_tag = soup.find('title')
        title = title_tag.get_text().strip() if title_tag else ""
        
        # Meta description
        desc_tag = soup.find('meta', attrs={'name': 'description'})
        description = desc_tag.get('content', '').strip() if desc_tag else ""
        
        # H1 tag
        h1_tag = soup.find('h1')
        h1 = h1_tag.get_text().strip() if h1_tag else ""
        
        # Canonical URL
        canonical_tag = soup.find('link', attrs={'rel': 'canonical'})
        canonical = canonical_tag.get('href', '') if canonical_tag else ""
        
        # Robots meta
        robots_tag = soup.find('meta', attrs={'name': 'robots'})
        robots = robots_tag.get('content', '') if robots_tag else ""
        noindex = 'Yes' if 'noindex' in robots.lower() else 'No'
        
        return {
            'title': title,
            'description': description,
            'h1': h1,
            'canonical': canonical,
            'noindex': noindex
        }
    except Exception as e:
        return {
            'title': f'Error: {str(e)}',
            'description': '',
            'h1': '',
            'canonical': '',
            'noindex': 'Error'
        }

def analyze_single_url(session, url):
    """Analyze one URL and return results"""
    try:
        response = session.get(url, timeout=REQUEST_TIMEOUT, allow_redirects=True)
        
        # Basic response info
        status_code = response.status_code
        final_url = response.url
        redirect_count = len(response.history)
        
        # Parse HTML and extract SEO data
        soup = BeautifulSoup(response.content, 'html.parser')
        seo_data = extract_seo_data(soup, final_url)
        
        return {
            'original_url': url,
            'final_url': final_url,
            'status_code': status_code,
            'redirect_count': redirect_count,
            'title': seo_data['title'],
            'description': seo_data['description'],
            'h1': seo_data['h1'],
            'canonical': seo_data['canonical'],
            'noindex': seo_data['noindex']
        }
        
    except Exception as e:
        return {
            'original_url': url,
            'final_url': url,
            'status_code': 'Error',
            'redirect_count': 0,
            'title': f'Error: {str(e)}',
            'description': '',
            'h1': '',
            'canonical': '',
            'noindex': 'Error'
        }

def get_urls_from_sitemap(session, sitemap_url):
    """Extract URLs from XML sitemap"""
    try:
        response = session.get(sitemap_url, timeout=REQUEST_TIMEOUT)
        response.raise_for_status()
        
        # Parse XML
        root = ET.fromstring(response.content)
        
        urls = []
        # Look for <loc> tags which contain URLs
        for element in root.iter():
            if element.tag.endswith('loc') and element.text:
                url = element.text.strip()
                if url.startswith('http'):
                    urls.append(url)
        
        return list(set(urls))  # Remove duplicates
        
    except Exception as e:
        print(f"Sitemap error: {e}")
        return []

def create_excel_file(data_list):
    """Create Excel file from data"""
    wb = Workbook()
    ws = wb.active
    ws.title = "SEO Analysis"
    
    # Headers
    headers = [
        'Original URL', 'Final URL', 'Status Code', 'Redirects',
        'Title', 'Meta Description', 'H1', 'Canonical URL', 'Noindex'
    ]
    
    # Style headers
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
    
    # Add data
    for row, item in enumerate(data_list, 2):
        ws.cell(row=row, column=1, value=item['original_url'])
        ws.cell(row=row, column=2, value=item['final_url'])
        ws.cell(row=row, column=3, value=str(item['status_code']))
        ws.cell(row=row, column=4, value=item['redirect_count'])
        ws.cell(row=row, column=5, value=item['title'])
        ws.cell(row=row, column=6, value=item['description'])
        ws.cell(row=row, column=7, value=item['h1'])
        ws.cell(row=row, column=8, value=item['canonical'])
        ws.cell(row=row, column=9, value=item['noindex'])
    
    # Adjust column widths
    column_widths = [40, 40, 12, 10, 50, 60, 40, 40, 10]
    for col, width in enumerate(column_widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = width
    
    # Save to memory
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

@app.route('/')
def home():
    """Basic API info"""
    return jsonify({
        'message': 'SEO Analyzer Pro API',
        'status': 'running',
        'endpoints': ['/api/analyze', '/api/download/<filename>']
    })

@app.route('/api/analyze', methods=['POST'])
def analyze_website():
    """Main analysis endpoint"""
    try:
        # Get sitemap URL from request
        data = request.get_json()
        sitemap_url = data.get('sitemap_url', '').strip()
        
        if not sitemap_url:
            return jsonify({'error': 'Sitemap URL is required'}), 400
        
        # Create session
        session = create_session()
        start_time = time.time()
        
        # Get URLs from sitemap
        urls = get_urls_from_sitemap(session, sitemap_url)
        
        if not urls:
            return jsonify({'error': 'No URLs found in sitemap'}), 400
        
        # Limit URLs for demo (remove for paid version)
        if len(urls) > 50:
            urls = urls[:50]
        
        # Analyze URLs using threading
        results = []
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            # Submit all tasks
            future_to_url = {
                executor.submit(analyze_single_url, session, url): url 
                for url in urls
            }
            
            # Collect results
            for future in as_completed(future_to_url):
                result = future.result()
                results.append(result)
        
        # Calculate stats
        total_urls = len(results)
        errors = sum(1 for r in results if 'Error' in str(r['status_code']))
        missing_titles = sum(1 for r in results if not r['title'].strip())
        missing_descriptions = sum(1 for r in results if not r['description'].strip())
        
        # Create Excel file
        excel_data = create_excel_file(results)
        
        # Generate filename
        domain = urlparse(sitemap_url).netloc.replace('.', '_')
        filename = f"{domain}_seo_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        # Store file (in production, use proper storage)
        if not hasattr(app, 'files'):
            app.files = {}
        app.files[filename] = excel_data
        
        analysis_time = time.time() - start_time
        
        return jsonify({
            'success': True,
            'totalUrls': total_urls,
            'analysisTime': f"{analysis_time:.1f} seconds",
            'downloadFilename': filename,
            'stats': {
                'processed': total_urls,
                'errors': errors,
                'warnings': missing_titles + missing_descriptions,
                'healthy': total_urls - errors
            },
            'issues': {
                'missingTitles': missing_titles,
                'missingDescriptions': missing_descriptions,
                'errors': errors
            },
            'categories': {
                'All Pages': total_urls
            }
        })
        
    except Exception as e:
        return jsonify({'error': f'Analysis failed: {str(e)}'}), 500

@app.route('/api/download/<filename>')
def download_file(filename):
    """Download Excel report"""
    try:
        if not hasattr(app, 'files') or filename not in app.files:
            return jsonify({'error': 'File not found'}), 404
        
        file_data = app.files[filename]
        file_data.seek(0)
        
        return send_file(
            file_data,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        return jsonify({'error': f'Download failed: {str(e)}'}), 500

if __name__ == '__main__':
    print("Starting SEO Analyzer Pro...")
    app.run(debug=True, host='0.0.0.0', port=5000)