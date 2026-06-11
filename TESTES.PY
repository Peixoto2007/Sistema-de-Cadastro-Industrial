from ollama import chat

print("ANTES", flush=True)

resposta = chat(
    model='qwen2.5:1.5b',
    messages=[
        {'role': 'user', 'content': 'Responda apenas OK'}
    ]
)

print("DEPOIS", flush=True)
print(resposta.message.content)