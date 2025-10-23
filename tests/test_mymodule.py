"""
Tests for MyModule

This module contains unit tests for the MyModule package.
"""

import pytest
import sys
from pathlib import Path

# Add src directory to Python path for testing
src_path = Path(__file__).parent.parent / "src"
sys.path.insert(0, str(src_path))

from MyModule import hello_world, Calculator


class TestHelloWorld:
    """Test cases for the hello_world function."""
    
    def test_hello_world_default(self):
        """Test hello_world with default parameter."""
        result = hello_world()
        assert result == "Hello, World!"
    
    def test_hello_world_with_name(self):
        """Test hello_world with custom name."""
        result = hello_world("Python")
        assert result == "Hello, Python!"
    
    def test_hello_world_empty_string(self):
        """Test hello_world with empty string."""
        result = hello_world("")
        assert result == "Hello, !"
    
    def test_hello_world_return_type(self):
        """Test that hello_world returns a string."""
        result = hello_world()
        assert isinstance(result, str)


class TestCalculator:
    """Test cases for the Calculator class."""
    
    def test_add(self):
        """Test addition operation."""
        assert Calculator.add(2, 3) == 5
        assert Calculator.add(-1, 1) == 0
        assert Calculator.add(0, 0) == 0
        assert Calculator.add(1.5, 2.5) == 4.0
    
    def test_subtract(self):
        """Test subtraction operation."""
        assert Calculator.subtract(5, 3) == 2
        assert Calculator.subtract(0, 5) == -5
        assert Calculator.subtract(10, 10) == 0
        assert Calculator.subtract(7.5, 2.5) == 5.0
    
    def test_multiply(self):
        """Test multiplication operation."""
        assert Calculator.multiply(3, 4) == 12
        assert Calculator.multiply(-2, 3) == -6
        assert Calculator.multiply(0, 100) == 0
        assert Calculator.multiply(2.5, 4) == 10.0
    
    def test_divide(self):
        """Test division operation."""
        assert Calculator.divide(10, 2) == 5.0
        assert Calculator.divide(7, 2) == 3.5
        assert Calculator.divide(-6, 3) == -2.0
        assert Calculator.divide(0, 5) == 0.0
    
    def test_divide_by_zero(self):
        """Test division by zero raises ValueError."""
        with pytest.raises(ValueError, match="Cannot divide by zero"):
            Calculator.divide(10, 0)
    
    def test_calculator_methods_are_static(self):
        """Test that calculator methods can be called without instantiation."""
        # These should work without creating a Calculator instance
        assert Calculator.add(1, 1) == 2
        assert Calculator.subtract(2, 1) == 1
        assert Calculator.multiply(2, 3) == 6
        assert Calculator.divide(6, 2) == 3.0


class TestModuleImports:
    """Test cases for module imports and metadata."""
    
    def test_module_imports(self):
        """Test that all expected functions and classes can be imported."""
        from MyModule import hello_world, Calculator, __version__, __author__, __email__
        
        # Check that imports work
        assert callable(hello_world)
        assert Calculator is not None
        
        # Check metadata
        assert isinstance(__version__, str)
        assert isinstance(__author__, str)
        assert isinstance(__email__, str)
    
    def test_module_all_exports(self):
        """Test that __all__ contains expected exports."""
        import MyModule
        expected_exports = ["hello_world", "Calculator"]
        assert hasattr(MyModule, "__all__")
        assert set(MyModule.__all__) == set(expected_exports)