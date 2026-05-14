from src.utils.font import normalize_font_name, linewidth_to_border_style


class TestNormalizeFontName:
    def test_none_returns_none(self):
        assert normalize_font_name(None) is None

    def test_empty_string_returns_none(self):
        assert normalize_font_name('') is None

    def test_bytes_returns_none(self):
        assert normalize_font_name(b'\x80\x81') is None

    def test_bytes_repr_string_returns_none(self):
        assert normalize_font_name("b'\\x80\\x81'") is None

    def test_subset_prefix_stripped(self):
        assert normalize_font_name('ABCDEF+MS-Gothic') == 'MS-Gothic'
        assert normalize_font_name('BCDEFG+MSMincho') == 'MSMincho'

    def test_no_prefix_returned_as_is(self):
        assert normalize_font_name('MS-Gothic') == 'MS-Gothic'
        assert normalize_font_name('MSGothic') == 'MSGothic'
        assert normalize_font_name('MSPGothic') == 'MSPGothic'
        assert normalize_font_name('MeiryoUI') == 'MeiryoUI'
        assert normalize_font_name('YuGothic') == 'YuGothic'
        assert normalize_font_name('MS-Mincho') == 'MS-Mincho'
        assert normalize_font_name('MSPMincho') == 'MSPMincho'
        assert normalize_font_name('YuMincho') == 'YuMincho'
        assert normalize_font_name('小塚明朝') == '小塚明朝'
        assert normalize_font_name('Arial') == 'Arial'


class TestLinewidthToBorderStyle:
    def test_thin(self):
        assert linewidth_to_border_style(0.0) == 'thin'
        assert linewidth_to_border_style(0.5) == 'thin'
        assert linewidth_to_border_style(1.0) == 'thin'

    def test_medium(self):
        assert linewidth_to_border_style(1.1) == 'medium'
        assert linewidth_to_border_style(2.0) == 'medium'

    def test_thick(self):
        assert linewidth_to_border_style(2.1) == 'thick'
        assert linewidth_to_border_style(5.0) == 'thick'
