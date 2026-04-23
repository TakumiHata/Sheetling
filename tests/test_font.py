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

    def test_gothic_variants_return_gothic(self):
        assert normalize_font_name('ABCDEF+MS-Gothic') == 'MS ゴシック'
        assert normalize_font_name('MS-Gothic') == 'MS ゴシック'
        assert normalize_font_name('MSGothic') == 'MS ゴシック'
        assert normalize_font_name('MSPGothic') == 'MS ゴシック'
        assert normalize_font_name('MeiryoUI') == 'MS ゴシック'
        assert normalize_font_name('YuGothic') == 'MS ゴシック'

    def test_mincho_variants_return_mincho(self):
        assert normalize_font_name('BCDEFG+MSMincho') == 'MS 明朝'
        assert normalize_font_name('MS-Mincho') == 'MS 明朝'
        assert normalize_font_name('MSPMincho') == 'MS 明朝'
        assert normalize_font_name('YuMincho') == 'MS 明朝'

    def test_japanese_mincho_name(self):
        assert normalize_font_name('小塚明朝') == 'MS 明朝'

    def test_unknown_font_defaults_to_mincho(self):
        assert normalize_font_name('Arial') == 'MS 明朝'


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
