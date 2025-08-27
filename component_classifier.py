import re
from typing import Dict, Tuple, Optional

class ComponentClassifier:
    """Component classification and default value assignment"""
    
    def __init__(self):
        # Component type patterns and their default values
        self.component_patterns = {
            'resistor': {
                'patterns': [r'^R\d+', r'RK\d+', r'RC\d+', r'RF\d+', r'RN\d+', r'resistor', r'抵抗'],
                'defaults': {
                    '品名': 'チップ抵抗',
                    '部品型番': 'SMD',
                    '実装/検査': '実装'
                }
            },
            'capacitor': {
                'patterns': [r'^C\d+', r'CC\d+', r'CG\d+', r'capacitor', r'コンデンサ'],
                'defaults': {
                    '品名': 'チップコンデンサ',
                    '部品型番': 'SMD',
                    '実装/検査': '実装'
                }
            },
            'inductor': {
                'patterns': [r'^L\d+', r'LK\d+', r'inductor', r'インダクタ'],
                'defaults': {
                    '品名': 'チップインダクタ',
                    '部品型番': 'SMD',
                    '実装/検査': '実装'
                }
            },
            'diode': {
                'patterns': [r'^D\d+', r'BAS\d+', r'BAT\d+', r'diode', r'ダイオード'],
                'defaults': {
                    '品名': 'ダイオード',
                    '部品型番': 'SMD',
                    '実装/検査': '実装'
                }
            },
            'transistor': {
                'patterns': [r'^Q\d+', r'^T\d+', r'BSS\d+', r'BC\d+', r'transistor', r'トランジスタ'],
                'defaults': {
                    '品名': 'トランジスタ',
                    '部品型番': 'SMD',
                    '実装/検査': '実装'
                }
            },
            'ic': {
                'patterns': [r'^U\d+', r'^IC\d+', r'ATmega', r'PIC\d+', r'STM32', r'LM\d+', r'TL\d+'],
                'defaults': {
                    '品名': 'IC',
                    '部品型番': 'SMD',
                    '実装/検査': '実装'
                }
            },
            'connector': {
                'patterns': [r'^J\d+', r'^CN\d+', r'^P\d+', r'connector', r'コネクタ'],
                'defaults': {
                    '品名': 'コネクタ',
                    '部品型番': 'DIP',
                    '実装/検査': '実装'
                }
            },
            'crystal': {
                'patterns': [r'^X\d+', r'^Y\d+', r'crystal', r'水晶', r'クリスタル'],
                'defaults': {
                    '品名': '水晶振動子',
                    '部品型番': 'SMD',
                    '実装/検査': '実装'
                }
            }
        }
        
        # Manufacturer-specific defaults
        self.manufacturer_defaults = {
            'KOA': {'品名': 'チップ抵抗'},
            'Murata': {'品名': 'チップコンデンサ'},
            'TDK': {'品名': 'チップコンデンサ'},
            'Panasonic': {'品名': 'チップコンデンサ'},
            'Yageo': {'品名': 'チップ抵抗'},
            'Vishay': {'品名': 'チップ抵抗'},
        }
    
    def classify_component(self, part_number: str, manufacturer: str = '', 
                          ref_designator: str = '') -> Dict[str, str]:
        """Classify component and return default values"""
        defaults = {
            '品名': '',
            '部品型番': 'SMD',  # Default to SMD
            '実装/検査': '実装'
        }
        
        # Check reference designator first (most reliable)
        if ref_designator:
            for component_type, info in self.component_patterns.items():
                for pattern in info['patterns']:
                    if re.search(pattern, ref_designator, re.IGNORECASE):
                        defaults.update(info['defaults'])
                        return defaults
        
        # Check part number patterns
        if part_number:
            for component_type, info in self.component_patterns.items():
                for pattern in info['patterns']:
                    if re.search(pattern, part_number, re.IGNORECASE):
                        defaults.update(info['defaults'])
                        return defaults
        
        # Check manufacturer-specific defaults
        if manufacturer in self.manufacturer_defaults:
            defaults.update(self.manufacturer_defaults[manufacturer])
        
        return defaults
    
    def detect_package_type(self, part_number: str, manufacturer: str = '') -> str:
        """Detect package type (SMD/DIP/BGA)"""
        part_upper = part_number.upper()
        
        # BGA patterns
        bga_patterns = ['BGA', 'FBGA', 'UBGA', 'CBGA']
        for pattern in bga_patterns:
            if pattern in part_upper:
                return '組込(BGA他)'
        
        # DIP patterns
        dip_patterns = ['DIP', 'PDIP', 'SOIC', 'SOP', 'SSOP', 'TSSOP', 'QFP', 'LQFP', 'TQFP']
        for pattern in dip_patterns:
            if pattern in part_upper:
                return 'DIP'
        
        # Size-based SMD detection (common chip sizes)
        smd_sizes = ['0201', '0402', '0603', '0805', '1206', '1210', '2010', '2512']
        for size in smd_sizes:
            if size in part_upper:
                return 'SMD'
        
        # Default to SMD for modern components
        return 'SMD'
    
    def enhance_component_data(self, row_data: Dict) -> Dict:
        """Enhance component data with defaults and classifications"""
        part_number = row_data.get('電子部品型番', '')
        manufacturer = row_data.get('メーカー', '')
        ref_designator = row_data.get('配置記号', '')
        
        # Get component classification defaults
        defaults = self.classify_component(part_number, manufacturer, ref_designator)
        
        # Apply defaults only if fields are empty
        for key, default_value in defaults.items():
            if not row_data.get(key, '').strip():
                row_data[key] = default_value
        
        # Detect and set package type
        if not row_data.get('部品型番', '').strip():
            row_data['部品型番'] = self.detect_package_type(part_number, manufacturer)
        
        return row_data
    
    def get_component_category(self, part_number: str, ref_designator: str = '') -> str:
        """Get component category for P-BAN.com classification"""
        categories = {
            'resistor': '抵抗器',
            'capacitor': 'コンデンサ',
            'inductor': 'インダクタ',
            'diode': 'ダイオード',
            'transistor': 'トランジスタ',
            'ic': 'IC',
            'connector': 'コネクタ',
            'crystal': '水晶振動子'
        }
        
        # Check reference designator first
        if ref_designator:
            for component_type, info in self.component_patterns.items():
                for pattern in info['patterns']:
                    if re.search(pattern, ref_designator, re.IGNORECASE):
                        return categories.get(component_type, '')
        
        # Check part number
        if part_number:
            for component_type, info in self.component_patterns.items():
                for pattern in info['patterns']:
                    if re.search(pattern, part_number, re.IGNORECASE):
                        return categories.get(component_type, '')
        
        return ''

def test_classifier():
    """Test the component classifier"""
    classifier = ComponentClassifier()
    
    test_cases = [
        {'part': 'RK73B2ATTD1002F', 'ref': 'R1,R5', 'mfr': 'KOA'},
        {'part': 'C1608X7R1H104K080AA', 'ref': 'C1,C2', 'mfr': 'TDK'},
        {'part': 'ATmega328P-PU', 'ref': 'U1', 'mfr': 'Microchip'},
        {'part': 'BAT54S', 'ref': 'D1', 'mfr': 'Vishay'},
    ]
    
    for case in test_cases:
        result = classifier.classify_component(
            case['part'], case['mfr'], case['ref']
        )
        print(f"Part: {case['part']}, Ref: {case['ref']} -> {result}")

if __name__ == "__main__":
    test_classifier()