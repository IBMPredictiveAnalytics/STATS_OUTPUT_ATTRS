#/***********************************************************************
# * Licensed Materials - Property of IBM 
# *
# * IBM SPSS Products: Statistics Common
# *
# * (C) Copyright IBM Corp. 1989, 2014
# *
# * US Government Users Restricted Rights - Use, duplication or disclosure
# * restricted by GSA ADP Schedule Contract with IBM Corp. 
# ************************************************************************/


"""STATS OUTPUT ATTRS extension command"""

__author__ =  'JKP'
__version__=  '1.0.1'

# history
# 07-nov-2011 original version
# 30-jan-2012 work around encoding bug in SetHeaderText and SetFooterText apis


helptext = """
This command sets options for printed or exported output.  It has no effect on
the display in the Viewer.

STATS OUTPUT ATTRS HEADER="header text" FOOTER = "footer text"
MARGINS=left right top bottom ORIENTATION={PORTRAIT | LANDSCAPE}
PAGENUMBER=number ITEMSPACING=number
[/HELP]

All keywords are optional.

Example:
STATS OUTPUT ATTRS HEADER="My Header Text" ORIENTATION=LANDSCAPE.

HEADER and FOOTER text specify text for the page header and footer for printed
or exported output.  It can be a single quoted string of text or multiple strings, e.g.
HEADER="first line of header" "second line".  You can include the same special
notation as appears in the Page Attributes dialog box.  For example, including
&[Date] inserts the current date in the header.  See the dialog box help for details.

The text can be plain text or html.  If the first item starts with <html, it is considered
to be html, and items are joined without.
Be sure to close the last line with </html>.
If it is not plain text, multiple strings are joined without a any break or other separator.
Use appropriate markup such as <br> for html to specify line breaking.

MARGINS specify the page margins.  If this keyword is used, there must be 
four numbers for the left, right, top, and bottom margins, respectively.  
Margins are measured in points (72 pts = 1 inch), and the values must be 
integers.  To leave a value at its current setting, enter a negative integer.
Margins refer to the page in Portrait orientation.

ORIENTATION can set the page orientation to be portrait or landscape.

PAGENUMBER sets the starting page number.

ITEMSPACING sets the inter-item spacing distance.  It is measured in points and
must be a nonnegative integer.
"""

#CHARTSIZE={ASIS | FULLPAGE | HALFPAGE |QUARTERPAGE}
#CHARTSIZE controls the size of printed or exported charts.  ASIS retains the Viewer
#size.  The other keywords set the size to full, half, or quarter page.

import spss, spssaux, SpssClient
from extension import Template, Syntax, processcmd

def outputAttrs(header=None, footer=None, margins=None,
    orientation=None, pagenumber=None, itemspacing=None, chartsize=None):
    """Set printing properties for designated Viewer window or globally"""
    
        # debugging
    # makes debug apply only to the current thread
    #try:
        #import wingdbstub
        #if wingdbstub.debugger != None:
            #import time
            #wingdbstub.debugger.StopDebug()
            #time.sleep(2)
            #wingdbstub.debugger.StartDebug()
        #import thread
        #wingdbstub.debugger.SetDebugThreads({thread.get_ident(): 1}, default_policy=0)
        ## for V19 use
        #SpssClient._heartBeat(False)
    #except:
        #pass
    
    if margins and len(margins) != 4:
        raise ValueError(_("""MARGINS must specify four values (left, right, top, bottom in points)"""))
    try:
        SpssClient.StartClient()
        desout = SpssClient.GetDesignatedOutputDoc()  # empty window gets created if none

        # header and footer text
        # headers and footers are always handled as html, so a <br> causes a line break
        # text must be in UTF-8 encoding
        for item, api in zip([header, footer], [desout.SetHeaderText, desout.SetFooterText]):
            if item:
                item = [line for line in item if line]
                item = "<br>".join(item)
                item = item.encode("UTF-8", "replace")
                api(item)
                
        # page margins
        if margins:
            opts = [SpssClient.PrintOptions.LeftMargin, SpssClient.PrintOptions.RightMargin,
                SpssClient.PrintOptions.TopMargin, SpssClient.PrintOptions.BottomMargin]
            for i in range(4):
                if margins[i] >= 0:
                    desout.SetPrintOptions(opts[i], str(margins[i]))
            
        # page orientation
        if orientation:
            desout.SetPrintOptions(SpssClient.PrintOptions.Orientation, 
                orientation == "portrait" and "1" or "2")
        
        # starting page number
        if pagenumber:
            desout.SetPrintOptions(SpssClient.PrintOptions.StartingPageNumber, str(pagenumber))
            
        # inter-item spacing
        if itemspacing:
            desout.SetPrintOptions(SpssClient.PrintOptions.SpaceBetweenItems, str(itemspacing))
            
        # chart size
        # feature removed as api does not work
        #if chartsize:
            #parm = ["asis", "fullpage", "halfpage", "quarterpage"].index(chartsize)  # already validated
            #desout.SetPrintOptions(SpssClient.PrintOptions.PrintedChartSize, str(parm))

    finally:
        SpssClient.StopClient()

def Run(args):
    """Execute the STATS OUTPUT ATTRS extension command"""

    args = args[args.keys()[0]]

    oobj = Syntax([
        Template("HEADER", subc="",  ktype="literal", var="header", islist=True),
        Template("FOOTER", subc="",  ktype="literal", var="footer", islist=True),
        Template("MARGINS", subc="", ktype="int", var="margins", islist=True),
        Template("ORIENTATION", subc="", ktype="str", var="orientation",
            vallist=["portrait","landscape"]),
        Template("PAGENUMBER", subc="", ktype="int", var="pagenumber"),
        Template("ITEMSPACING", subc="", ktype="int", var="itemspacing", vallist=[0]),
        Template("HELP", subc="", ktype="bool")])
    
        # Template("CHARTSIZE", subc="", ktype="str", var="chartsize",
        # vallist=["asis", "fullpage", "halfpage", "quarterpage"]),
    
    #enable localization
    global _
    try:
        _("---")
    except:
        def _(msg):
            return msg
    # A HELP subcommand overrides all else
    if args.has_key("HELP"):
        #print helptext
        helper()
    else:
        processcmd(oobj, args, outputAttrs)

def helper():
    """open html help in default browser window
    
    The location is computed from the current module name"""
    
    import webbrowser, os.path
    
    path = os.path.splitext(__file__)[0]
    helpspec = "file://" + path + os.path.sep + \
         "markdown.html"
    
    # webbrowser.open seems not to work well
    browser = webbrowser.get()
    if not browser.open_new(helpspec):
        print("Help file not found:" + helpspec)
try:    #override
    from extension import helper
except:
    pass