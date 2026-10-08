from pathlib import Path

from generate_labels import _build_parser


def test_bare_command_uses_example_profile_by_default():
    args = _build_parser().parse_args([])

    assert args.input_xlsx is None
    assert args.output_docx is None
    assert args.profile == Path("profiles/example.yaml")
    assert args.template == Path("templates/template.docx")
    assert args.workbook_dir == Path("workbook")
