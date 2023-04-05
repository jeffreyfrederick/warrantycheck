# Warranty Status Check
 <img src="/images/logo.svg" height="100" width="100"> 
<p>Automates checking the warranty status of Dell products from an Excel file utilizing ayncio and the Playwright API.</p>

<h2>Features</h2>
<ul>
 <li>Reads device information from an input Excel file</li>
 <li>Scrapes Dell's warranty page for warranty expiration dates</li>
 <li>Handles anti-bot measures and popups on the warranty page</li>
 <li>Outputs the collected data to a new Excel file>/li>
</ul>
 
<h2>Dependencies</h2>
<ul>
 <li>Python 3.7+</li>
 <li>playwright (for browser automation)</li>
 <li>pandas (for data manipulation and Excel file I/O)</li>
 <li>httpx (for asynchronous HTTP client)</li>
 <li>aiolimiter (for rate limiting)</li>
 <li>tenacity (for retry logic)</li>
 <li>openpyxl (for writing Excel files)</li>
</ul>

<p>To install the required dependencies, run:</p>
<pre><code>pip install playwright pandas httpx aiolimiter tenacity openpyxl</code></pre>

<h2>Usage</h2>
<ol>
    <li>Prepare an input Excel file named <code>input.xlsx</code> in the following format:</li>
</ol>
<table>
    <thead>
        <tr>
            <th>Unique #</th>
            <th>Serial #</th>
            <th>OEM</th>
            <th>Model</th>
            <th>Location</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td>1</td>
            <td>ABC123</td>
            <td>Dell</td>
            <td>XPS 13</td>
            <td>NY</td>
        </tr>
    </tbody>
</table>
<ol start="2">
    <li>Place the input file in the same directory as the script.</li>
    <li>Run the script with:</li>
</ol>
<pre><code>python dell_warranty_checker.py</code></pre>
<ol start="4">
    <li>After the script finishes, you will find an output file named <code>output.xlsx</code> with the following format:</li>
</ol>
<table>
    <thead>
        <tr>
            <th>Unique #</th>
            <th>Serial #</th>
            <th>OEM</th>
            <th>Model</th>
            <th>Expiration</th>
            <th>Location</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td>1</td>
            <td>ABC123</td>
            <td>Dell</td>
            <td>XPS 13</td>
            <td>2023-12-31</td>
            <td>NY</td>
        </tr>
    </tbody>
</table>

<h2>Note</h2>
<p>This script uses an asynchronous approach to scrape warranty information and rate limits the tasks to avoid overloading your system. However, it is possible that Dell's website could change, which may affect the functionality of the script. Please use this tool responsibly and respect the website's terms of service.</p>

<h2>License</h2>
<p>This project is licensed under the <a href="LICENSE">MIT License</a>.</p>
