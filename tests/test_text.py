from src.utils.text import has_japanese, join_word_texts, split_by_horizontal_gap


class TestHasJapanese:
    def test_hiragana(self):
        assert has_japanese('あいう') is True

    def test_katakana(self):
        assert has_japanese('アイウ') is True

    def test_kanji(self):
        assert has_japanese('漢字') is True

    def test_fullwidth(self):
        assert has_japanese('ＡＢＣ') is True

    def test_ascii_only(self):
        assert has_japanese('Hello World') is False

    def test_empty(self):
        assert has_japanese('') is False

    def test_mixed(self):
        assert has_japanese('Hello世界') is True


class TestJoinWordTexts:
    def test_japanese_no_space(self):
        assert join_word_texts(['東京', '都']) == '東京都'

    def test_english_with_space(self):
        assert join_word_texts(['Hello', 'World']) == 'Hello World'

    def test_mixed_no_space(self):
        assert join_word_texts(['Hello', '世界']) == 'Hello世界'

    def test_single_word(self):
        assert join_word_texts(['test']) == 'test'

    def test_empty_list(self):
        assert join_word_texts([]) == ''

    def test_strips_empty_english(self):
        assert join_word_texts(['Hello', ' ', 'World']) == 'Hello World'


class TestSplitByHorizontalGap:
    def test_single_word(self):
        words = [{'x0': 0, 'x1': 10, 'font_size': 10}]
        result = split_by_horizontal_gap(words)
        assert len(result) == 1
        assert len(result[0]) == 1

    def test_no_gap(self):
        words = [
            {'x0': 0, 'x1': 10, 'font_size': 10},
            {'x0': 11, 'x1': 20, 'font_size': 10},
        ]
        result = split_by_horizontal_gap(words)
        assert len(result) == 1

    def test_large_gap_splits(self):
        words = [
            {'x0': 0, 'x1': 10, 'font_size': 10},
            {'x0': 50, 'x1': 60, 'font_size': 10},
        ]
        result = split_by_horizontal_gap(words)
        assert len(result) == 2

    def test_empty_list(self):
        result = split_by_horizontal_gap([])
        assert result == [[]]

    def test_custom_gap_factor(self):
        words = [
            {'x0': 0, 'x1': 10, 'font_size': 10},
            {'x0': 25, 'x1': 35, 'font_size': 10},
        ]
        assert len(split_by_horizontal_gap(words, gap_factor=1.0)) == 2
        assert len(split_by_horizontal_gap(words, gap_factor=3.0)) == 1
