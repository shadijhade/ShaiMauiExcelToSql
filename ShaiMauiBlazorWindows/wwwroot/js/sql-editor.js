export function init(dotNetObject, textAreaId) {
    // Check if textarea exists
    const textarea = document.getElementById(textAreaId);
    if (!textarea) {
        console.error('Textarea with id', textAreaId, 'not found');
        return null;
    }
    
    // Get current theme
    const currentTheme = document.documentElement.getAttribute('data-bs-theme') || 'light';
    const editorTheme = currentTheme === 'dark' ? 'material-darker' : 'neat';
    
    const editor = CodeMirror.fromTextArea(textarea, {
        mode: 'text/x-sql',
        lineNumbers: true,
        theme: editorTheme,
        extraKeys: { "Ctrl-Space": "autocomplete" },
        hintOptions: {
            tables: {
                users: ["name", "score", "birthDate"],
                countries: ["name", "population", "size"]
            }
        }
    });

    editor.on('change', () => {
        const content = editor.getValue();
        dotNetObject.invokeMethodAsync('UpdateQuery', content);
    });

    // Listen for theme changes
    try {
        const observer = new MutationObserver(function(mutations) {
            mutations.forEach(function(mutation) {
                if (mutation.type === 'attributes' && mutation.attributeName === 'data-bs-theme') {
                    try {
                        const newTheme = document.documentElement.getAttribute('data-bs-theme') || 'light';
                        const newEditorTheme = newTheme === 'dark' ? 'material-darker' : 'neat';
                        if (editor && editor.setOption) {
                            editor.setOption('theme', newEditorTheme);
                        }
                    } catch (error) {
                        console.warn('Failed to update editor theme:', error);
                    }
                }
            });
        });

        observer.observe(document.documentElement, {
            attributes: true,
            attributeFilter: ['data-bs-theme']
        });
    } catch (error) {
        console.warn('Failed to set up theme observer:', error);
    }

    return editor;
}

export function setValue(editor, text) {
    if (editor && editor.setValue) {
        editor.setValue(text);
    }
}

export function dispose(editor) {
    if (editor && editor.toTextArea) {
        editor.toTextArea();
    }
}