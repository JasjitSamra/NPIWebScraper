from openpyxl import load_workbook
from lxml import html
import requests
# Set workbook
WB = load_workbook("Dummy.xlsx")
# Set worksheet and column for worksheet
WS = WB.get_sheet_by_name('Sheet1')
COLUMN = WS['A']
# Convert column list to cell values for each value within range of number of columns
COLUMN_LIST = [COLUMN[x].value for x in range(len(COLUMN))]
# Remove header and convert to string
COLUMN_FILTER = COLUMN_LIST[1:]
#Setting empty lists to append in for loop`
NAMES = []
OTHER_NAMES = []
MAILING_ADDRESSES = []
PRIMARY_ADDRESSES = []
TAXONOMIES = []
# Set Direct Xpath for each column
DXPATH_NAME = "//html/body/div[2]/div/div[2]/div[1]/div[2]/blockquote/p/text()"
DXPATH_MAILING_ADDRESS = "/html/body/div[2]/div/div[2]/div[3]/table/tr[6]/td[2]/text()"
DXPATH_PRIMARY_ADDRESS = "/html/body/div[2]/div/div[2]/div[3]/table/tr[7]/td[2]/text()"
DXPATH_TAXONOMY = "/html/body/div[2]/div/div[2]/div[3]/table/tr[10]/td[2]/table/tr/td[2]/text()"
DXPATH_404CHECK = "/html/body/div[2]/div[2]/div/div/h1/span/text()"
# Creating For loop
for NPI in COLUMN_FILTER:
    # Setting the html for each NPI in column A
    VAR_HTML = "https://npiregistry.cms.hhs.gov/registry/provider-view/" + str(NPI)
    #print(var_html)
    # Retrieve web page
    PAGE = requests.get(VAR_HTML)
    # Parse using html module, save results in tree var
    TREE = html.fromstring(PAGE.content)
    # set equivalent variables to test if site down for if statement
    ERRORCH = TREE.xpath(DXPATH_404CHECK)
    ERRORCH = "".join(ERRORCH)
    ERRORCH2 = '404'
    # Site is down values
    if ERRORCH == ERRORCH2:
        NAME = "CMS Deactivated NPI"
        MAILING_ADDRESS = 0
        PRIMARY_ADDRESS = 0
        TAXONOMY = 0
        NAMES.append(NAME)
        MAILING_ADDRESSES.append(MAILING_ADDRESS)
        PRIMARY_ADDRESSES.append(PRIMARY_ADDRESS)
        TAXONOMIES.append(TAXONOMY)
        print("Shut down site found")
    # Pull direct Xpath elements\text() to list
    else:
        NAME = TREE.xpath(DXPATH_NAME)
        MAILING_ADDRESS = TREE.xpath(DXPATH_MAILING_ADDRESS)
        PRIMARY_ADDRESS = TREE.xpath(DXPATH_PRIMARY_ADDRESS)
        TAXONOMY = TREE.xpath(DXPATH_TAXONOMY)
    #Clean pulled text
        NAME[:] = [n.strip('   \n                  ') for n in NAME]
        MAILING_ADDRESS[:] = [ma.strip('\n\t\t\t\t\t\t\t\t\n\t\t\t\t\t\t\t\t')
                              for ma in MAILING_ADDRESS]
        MAILING_ADDRESS[:] = [ma[:-(len(' \n\t\t\t\t\t\t\t\t\n\t\t\t\t\t\t\t\t\t95825-1369'))]
                              for ma in MAILING_ADDRESS]
        PRIMARY_ADDRESS[:] = [pa.strip('\n\t\t\t\t\t\t\t\t\n\t\t\t\t\t\t\t\t')
                              for pa in PRIMARY_ADDRESS]
        PRIMARY_ADDRESS[:] = [pa[:-(len(' \n\t\t\t\t\t\t\t\t\n\t\t\t\t\t\t\t\t\n\t\t\t\t\t\t\t\t\t95825-1369'))]
                              for pa in PRIMARY_ADDRESS]
        TAXONOMY[:] = [t[len("\n\t\t\t\t\t\t\n\t\t\t\t\t\t\n\t\t\t\t\t\t"):]
                       for t in TAXONOMY]
        SPL_WORD = "  - "
        TAXONOMY[:] = TAXONOMY[-1:]
        TAXONOMY[:] = [t.partition(SPL_WORD)[2] for t in TAXONOMY]
    #Append lists
        NAMES.append(NAME)
        MAILING_ADDRESSES.append(MAILING_ADDRESS[1])
        PRIMARY_ADDRESSES.append(PRIMARY_ADDRESS[1])
        TAXONOMIES.append(TAXONOMY)
#Check first 10 hits for each list
print(NAMES[:10])
print(MAILING_ADDRESSES[:10])
print(PRIMARY_ADDRESSES[:10])
print(TAXONOMIES[:10])
#Write Name to second Column
#column = 2
#row = 2
#for i in enumerate(names):
