from questions.models import QuestionBankEntry


def normalize_text(value):
    if value is None:
        return ''
    return str(value).strip()


def split_module_path(module_path):
    parts = [
        normalize_text(part)
        for part in str(module_path or '').replace('\\', '/').split('/')
        if normalize_text(part)
    ]
    while len(parts) < 3:
        parts.append('')
    return tuple(parts[:3])


def build_question_bank_hierarchy():
    hierarchy = {}
    rows = QuestionBankEntry.objects.exclude(course_code='').values_list(
        'course_code',
        'module_path',
        'workbook_title',
        'topic',
    )
    for course_code, module_path, workbook_title, topic in rows.iterator():
        subcourse, module, workbook_from_path = split_module_path(module_path)
        workbook = normalize_text(workbook_title) or workbook_from_path
        topic_value = normalize_text(topic)

        course_node = hierarchy.setdefault(course_code, {})
        subcourse_node = course_node.setdefault(subcourse, {})
        module_node = subcourse_node.setdefault(module, {})
        workbook_topics = module_node.setdefault(workbook, set())
        if topic_value:
            workbook_topics.add(topic_value)

    serialized = {}
    for course_code, subcourses in sorted(hierarchy.items(), key=lambda item: item[0].lower()):
        serialized[course_code] = {}
        for subcourse, modules in sorted(subcourses.items(), key=lambda item: item[0].lower()):
            serialized[course_code][subcourse] = {}
            for module, workbooks in sorted(modules.items(), key=lambda item: item[0].lower()):
                serialized[course_code][subcourse][module] = {}
                for workbook, topics in sorted(workbooks.items(), key=lambda item: item[0].lower()):
                    serialized[course_code][subcourse][module][workbook] = sorted(topics, key=str.lower)
    return serialized


def hierarchy_choices(hierarchy, course_code, selected_subcourses=(), selected_modules=(), selected_workbooks=()):
    course_tree = hierarchy.get(course_code, {})
    selected_subcourses = set(selected_subcourses)
    selected_modules = set(selected_modules)
    selected_workbooks = set(selected_workbooks)

    subcourses = sorted(course_tree.keys(), key=str.lower)
    modules = set()
    workbooks = set()
    topics = set()

    for subcourse, modules_tree in course_tree.items():
        if selected_subcourses and subcourse not in selected_subcourses:
            continue
        for module, workbooks_tree in modules_tree.items():
            if module:
                modules.add(module)
            if selected_modules and module not in selected_modules:
                continue
            for workbook, workbook_topics in workbooks_tree.items():
                if workbook:
                    workbooks.add(workbook)
                if selected_workbooks and workbook not in selected_workbooks:
                    continue
                topics.update(topic for topic in workbook_topics if topic)

    return {
        'subcourses': sorted(subcourses, key=str.lower),
        'modules': sorted(modules, key=str.lower),
        'workbooks': sorted(workbooks, key=str.lower),
        'topics': sorted(topics, key=str.lower),
    }
