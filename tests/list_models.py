import os
import google.generativeai as genai

def force_load_env():
    env_path = ".env"
    if os.path.exists(env_path):
        with open(env_path, "r") as f:
            for line in f:
                if "=" in line and not line.startswith("#"):
                    k, v = line.strip().split("=", 1)
                    os.environ[k] = v.strip("'\"")

def list_models():
    force_load_env()
    api_key = os.getenv("GEMINI_API_KEY")
    genai.configure(api_key=api_key)
    print("利用可能なモデル一覧:")
    for m in genai.list_models():
        if 'generateContent' in m.supported_generation_methods:
            print(m.name)

if __name__ == "__main__":
    list_models()
