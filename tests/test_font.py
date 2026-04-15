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

    def test_subset_prefix_removed(self):
        assert normalize_font_name('ABCDEF+MS-Gothic') == 'MS Gothic'

    def test_hyphen_to_space(self):
        assert normalize_font_name('MS-Gothic') == 'MS Gothic'

    def test_alias_resolution(self):
        assert normalize_font_name('MSGothic') == 'MS Gothic'
        assert normalize_font_name('MSPGothic') == 'MS PGothic'
        assert normalize_font_name('MeiryoUI') == 'Meiryo UI'
        assert normalize_font_name('YuGothic') == 'Yu Gothic'

    def test_unknown_font_passthrough(self):
        assert normalize_font_name('Arial') == 'Arial'

    def test_subset_plus_alias(self):
        assert normalize_font_name('BCDEFG+MSMincho') == 'MS Mincho'


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
