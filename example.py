import jinja2


from template_pptx_jinja.render import PPTXRendering


def main():
    input_path = 'example/template.pptx'

    model = {
        "name": "John",
        "number": 3,
        "step": [
            {
                "name": "analysis"
            },
            {
                "name": "design"
            },
            {
                "name": "production"
            }
        ]
    }
    pictures = {
        "example/model.jpg": "example/image.jpg"
    }

    data = {
        'model': model,
        'pictures': pictures
    }

    def plural(input, word_ending):
        return word_ending if input > 0 else ''

    jinja2_env = jinja2.Environment()
    jinja2_env.filters['plural'] = plural

    output_path = 'example/presentation_generated.pptx'
    rendering = PPTXRendering(input_path, data, output_path, jinja2_env)
    message = rendering.process()
    print(message)

if __name__ == '__main__':
    main()
