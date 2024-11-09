import xlwings as xw
from xlwings.utils import hex_to_rgb
import logging

# Setup logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

def create_ribbon():
    """Create a custom ribbon with buttons"""
    logger.debug("Creating ribbon...")
    ribbon_xml = '''
    <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="ribbon_loaded">
        <ribbon startFromScratch="false">
            <tabs>
                <tab id="CustomTab" label="My Custom Tab" insertAfterMso="TabHome">
                    <group id="CustomGroup" label="My Tools">
                        <button id="Button1" 
                                label="Sample Button"
                                size="large"
                                onAction="sample_button_click"
                                imageMso="HappyFace"/>
                    </group>
                </tab>
            </tabs>
        </ribbon>
    </customUI>
    '''
    logger.debug(f"Ribbon XML: {ribbon_xml}")
    return ribbon_xml

@xw.func
def ribbon_loaded(ribbon):
    """Callback when the ribbon is loaded"""
    logger.info("Ribbon loaded successfully!")
    return None

@xw.func
def sample_button_click(event):
    """Function that runs when the sample button is clicked"""
    try:
        logger.info("Button click detected")
        wb = xw.Book.caller()
        wb.sheets[0].range('A1').value = "Button clicked!"
        logger.info("Button click processed successfully!")
    except Exception as e:
        logger.error(f"Error in button click: {str(e)}")
        raise

if __name__ == '__main__':
    logger.info("Starting xlwings server...")
    try:
        xw.serve()
    except Exception as e:
        logger.error(f"Server error: {str(e)}") 