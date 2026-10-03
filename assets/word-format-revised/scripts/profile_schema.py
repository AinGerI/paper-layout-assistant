"""Small strict validator for the implemented profile schema (no extra dependency)."""
import math
import re


def validate_profile(data):
    def mapping(value, allowed, path):
        if value is None:
            return {}
        if not isinstance(value, dict):
            raise ValueError(f'{path} must be a mapping')
        unknown = set(value) - set(allowed.split())
        if unknown:
            raise ValueError(f'{path}: unsupported keys {sorted(unknown)}')
        return value

    def check_values(node, path):
        for key, value in node.items():
            if value is None:
                continue
            where = f'{path}.{key}'
            if key in {'enabled', 'repaginate', 'bold', 'italic', 'underline', 'keep_with_next',
                       'keep_together', 'page_break_before', 'widow_control', 'autofit',
                       'allow_break_across_pages', 'repeat_header', 'lock_aspect_ratio', 'resize_floating_shapes'}:
                if not isinstance(value, bool):
                    raise ValueError(f'{where} must be true/false/null')
            elif key.endswith(('_pt', '_cm', '_chars', '_paragraphs')):
                if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
                    raise ValueError(f'{where} must be a finite number')
                if value < 0 and 'indent' not in key:
                    raise ValueError(f'{where} must be nonnegative')

    def enum(node, key, choices, path):
        if node.get(key) is not None and node[key] not in choices.split():
            raise ValueError(f'{path}.{key}: expected one of {choices}')

    def font(node, path):
        node = mapping(node, 'name east_asia chinese latin western size_pt bold italic underline', path)
        check_values(node, path)
        for key in ('name', 'east_asia', 'chinese', 'latin', 'western'):
            if node.get(key) is not None and not isinstance(node[key], str):
                raise ValueError(f'{path}.{key} must be text')

    def paragraph(node, path):
        node = mapping(node, 'alignment space_before_pt space_after_pt left_indent_pt right_indent_pt first_line_indent_chars first_line_indent_pt line_spacing keep_with_next keep_together page_break_before widow_control', path)
        check_values(node, path)
        enum(node, 'alignment', 'left center right justify distributed', path)
        spacing = node.get('line_spacing')
        if spacing is None:
            return
        if isinstance(spacing, dict):
            mapping(spacing, 'rule value', path + '.line_spacing')
            enum(spacing, 'rule', 'single 1.5 one_point_five double exact at_least atleast multiple multi', path)
            if spacing.get('rule') in ('exact', 'at_least', 'atleast', 'multiple', 'multi'):
                value = spacing.get('value')
                if value is not None and (isinstance(value, bool) or not isinstance(value, (int,float)) or not math.isfinite(value) or value <= 0):
                    raise ValueError(f'{path}.line_spacing.value must be positive or null')
        elif isinstance(spacing, (int, float)) and not isinstance(spacing, bool):
            if not math.isfinite(spacing) or spacing <= 0:
                raise ValueError(f'{path}.line_spacing must be positive')
        elif spacing not in ('single', '1', '1.0', '1.5', 'one_point_five', 'double', '2', '2.0'):
            raise ValueError(f'{path}.line_spacing is invalid')

    mapping(data, 'schema_version name source_template source_requirements extraction_notes document page roles table figure detection', 'profile')
    check_values(mapping(data.get('document'), 'repaginate', 'document'), 'document')
    page = mapping(data.get('page'), 'enabled orientation top_margin_cm bottom_margin_cm left_margin_cm right_margin_cm header_distance_cm footer_distance_cm gutter_cm paper_width_cm paper_height_cm', 'page')
    check_values(page, 'page')
    enum(page, 'orientation', 'portrait landscape', 'page')
    roles = data.get('roles') or {}
    valid_roles = 'title body equation_block figure_paragraph figure_caption table_caption abstract_heading keywords references_heading table_text empty ' + ' '.join(f'heading_{i}' for i in range(1,10))
    mapping(roles, valid_roles, 'roles')
    for role, spec in roles.items():
        spec = mapping(spec, 'enabled font paragraph', 'roles.' + role)
        check_values(spec, role)
        font(spec.get('font'), role + '.font')
        paragraph(spec.get('paragraph'), role + '.paragraph')
    table = mapping(data.get('table'), 'enabled alignment autofit preferred_width_cm font paragraph cell row border', 'table')
    check_values(table, 'table')
    enum(table, 'alignment', 'left center right', 'table')
    font(table.get('font'), 'table.font')
    paragraph(table.get('paragraph'), 'table.paragraph')
    cell = mapping(table.get('cell'), 'vertical_alignment margin_top_pt margin_bottom_pt margin_left_pt margin_right_pt', 'table.cell')
    check_values(cell, 'table.cell')
    enum(cell, 'vertical_alignment', 'top center bottom', 'table.cell')
    row = mapping(table.get('row'), 'height_rule height_pt allow_break_across_pages repeat_header', 'table.row')
    check_values(row, 'table.row')
    enum(row, 'height_rule', 'retain auto at_least atleast exact', 'table.row')
    border = mapping(table.get('border'), 'type style width_pt top_pt header_bottom_pt bottom_pt', 'table.border')
    check_values(border, 'table.border')
    enum(border, 'type', 'retain none grid full_grid three_line', 'table.border')
    enum(border, 'style', 'single double none', 'table.border')
    figure = mapping(data.get('figure'), 'enabled paragraph_alignment max_width_cm width_cm lock_aspect_ratio resize_floating_shapes', 'figure')
    check_values(figure, 'figure')
    enum(figure, 'paragraph_alignment', 'center', 'figure')
    detection = mapping(data.get('detection'), 'title_scan_paragraphs title_max_chars title_min_pt figure_caption_patterns table_caption_patterns paragraph_overrides regex_overrides', 'detection')
    check_values(detection, 'detection')
    for key in ('figure_caption_patterns', 'table_caption_patterns'):
        if not isinstance(detection.get(key, []), list):
            raise ValueError(f'detection.{key} must be a list')
        for pattern in detection.get(key, []):
            re.compile(pattern)
    for item in detection.get('regex_overrides') or []:
        mapping(item, 'pattern role', 'regex override')
        re.compile(item['pattern'])
        if item['role'] not in valid_roles.split():
            raise ValueError('Unknown role in regex override')
    for key, role in (detection.get('paragraph_overrides') or {}).items():
        if int(key) < 1 or role not in valid_roles.split():
            raise ValueError('Invalid paragraph override')
