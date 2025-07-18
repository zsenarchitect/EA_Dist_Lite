# -*- coding: utf-8 -*-

__dream__ = u'''
Despite everything, I still believe in the goodness of people.

The core concept of EnneadTab is to help people.
It is against the spirit of open source to restrict usage to certain group of people.
It wants to help and inspire more people to be creative and build their own tools.
The richness comes from the service I provided to others, not from owning a asset.
I work for the smile on people's face, and nothing else.
Don't you agree? 
Don't be a owner, be a giver.

Please spread the word and help me make this happen.
You are better than what is asked from you, despite everything.

Have a nice day, my friend, good and kind.

Sen.Z


尽管发生了一切，我仍相信人的善良。

EnneadTab的本质是帮助人们。
开源的精神不应使其限制于某些人的使用。
它想激励更多的人变得有创造力并构建自己的工具。
价值来自为他人提供的服务,而不是拥有资产。
我为人们的笑容而工作,除此之外别无他求。
你同意吗?
不要成为拥有者,要做给予者。

请传播这个想法并帮助我实现这个目标。
你好过你被迫成为的样子,不论发生什么。

祝你有美好的一天,我的朋友,善良和美好。

森.Z
'''



__package_name__ = "EnneadTab"
__version__  = "3.0"
"""
EnneadTab Package Information

This package is designed to be compatible with both IronPython 2.7 and Python 3.x environments.
All modules handle encoding properly without needing to manipulate sys.setdefaultencoding.

Note: Do not use reload(sys) or sys.setdefaultencoding('utf-8') as these can cause 
encoding failures in IronPython environments.
"""
import os
import sys
import traceback
import io

__py3_marker__ = "#!python3"



def get_module_files():
    """Get all Python module files in the current directory.
    
    Scans the package directory and identifies all Python modules that should be
    imported, excluding the __init__.py file itself.
    
    Returns:
        set: A set of strings containing the filenames of all .py files in the
            current directory that aren't __init__.py.
    
    Example:
        >>> get_module_files()
        {'REVIT.py', 'PDF.py', 'RHINO.py'}
    """
    return {
        module for module in os.listdir(os.path.dirname(__file__))
        if module.endswith('.py') and module != '__init__.py'
    }

def import_special_modules(module_name):
    """Handle special module imports (RHINO, REVIT).
    
    Some modules require special handling during import due to their dependencies
    or initialization requirements. This function handles those cases separately.
    
    Args:
        module_name (str): The name of the module to potentially import.
            Expected to be either "RHINO" or "REVIT".
    
    Returns:
        bool: True if the module was handled as a special case (regardless of
            whether the import succeeded), False if the module wasn't identified
            as needing special handling.
    
    Note:
        Special module import failures are silently ignored to prevent blocking
        the initialization of other modules.
    """
    if module_name not in ["RHINO", "REVIT"]:
        return False
        
    try:
        __import__("{}.{}".format(__package_name__, module_name), fromlist=['*'])
    except Exception:
        pass  # Silently skip if special module import fails
    return True

def import_module(module_name):
    """Import a single module with error handling.
    
    Attempts to import a module while handling potential import errors and
    ensuring the module directory is in the Python path.
    
    Args:
        module_name (str): The name of the module to import. Can include the .py
            extension, which will be stripped before import.
    
    Note:
        If an import fails, the error will be printed to stdout only if #!python3 is not present in the script.
        Two attempts are made to format the error message:
        1. Using the full traceback
        2. Using just the exception string if traceback formatting fails
    """
    try:
        # Ensure module directory is in path for relative imports
        module_dir = os.path.dirname(__file__)
        if module_dir not in sys.path:
            sys.path.append(module_dir)
            
        # Import the module (strip .py extension if present)
        base_name = module_name[:-3] if module_name.endswith('.py') else module_name
        module_path = os.path.join(module_dir, module_name)
        
        # Check if script has python3 marker
        should_silent = False
        if os.path.exists(module_path):
            with io.open(module_path, 'r', encoding='utf-8') as f:
                first_line = f.readline().strip()
                if first_line == __py3_marker__:
                    should_silent = True
                    
        __import__("{}.{}".format(__package_name__, base_name), fromlist=['*'])
    except Exception as e:
        if not should_silent:
            try:
                print("Cannot import {} because\n\n{}".format(
                    module_name, traceback.format_exc()))
            except:
                print("Cannot import {} because\n\n{}".format(
                    module_name, str(e)))

def initialize_package():
    """Initialize the package by importing all modules.
    
    This function orchestrates the package initialization process by:
    1. Getting a list of all Python modules in the package
    2. Attempting to import special modules first
    3. Importing remaining modules
    
    The function handles both regular and special module imports, ensuring that
    all package components are properly initialized.
    
    Note:
        This function is automatically called when the package is imported.
        Special modules (RHINO, REVIT) are handled separately from regular modules
        due to their specific initialization requirements.
    """
    for module in get_module_files():
        if not import_special_modules(module):
            import_module(module)

# Execute package initialization
initialize_package()



def dream():
    print(__dream__)
    return __dream__


