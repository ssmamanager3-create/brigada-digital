import { useEffect, useState } from "react";
import { Routes, Route, Navigate, useNavigate } from "react-router-dom";
import * as XLSX from "xlsx";
import {
  ShieldCheck,
  Users,
  UserCheck,
  UserX,
  Monitor,
  Settings,
  Search,
  LogIn,
  LogOut,
  ExternalLink,
  Clock,
  Eye,
  EyeOff,
  LockKeyhole,
  Camera,
  FileSpreadsheet,
  Plus,
  Trash2,
  Save,
  Upload,
  MapPin,
  Navigation,
  Smartphone,
  Crosshair,
  Bell,
  Phone,
  AlertTriangle,
} from "lucide-react";

import "./App.css";

const dadosIniciais = [
  {
    id: 1,
    nome: "Carlos Henrique",
    matricula: "BRG001",
    pin: "123456",
    funcao: "Coordenador de Emergência",
    equipe: "Coordenação",
    turno: "1º Turno",
    unidade: "Unidade Industrial",
    status: "Ativo",
    presente: true,
    foto: "https://i.pravatar.cc/300?img=12",
  },
  {
    id: 2,
    nome: "Marcos Silva",
    matricula: "BRG002",
    pin: "234567",
    funcao: "Líder da Brigada de Emergência",
    equipe: "Liderança",
    turno: "1º Turno",
    unidade: "Unidade Industrial",
    status: "Ativo",
    presente: true,
    foto: "https://i.pravatar.cc/300?img=11",
  },
  {
    id: 3,
    nome: "João Santos",
    matricula: "BRG003",
    pin: "345678",
    funcao: "Brigadista de Emergência",
    equipe: "Brigadistas de Emergência",
    turno: "1º Turno",
    unidade: "Unidade Industrial",
    status: "Ativo",
    presente: true,
    foto: "https://i.pravatar.cc/300?img=13",
  },
  {
    id: 4,
    nome: "Juliana Lima",
    matricula: "BRG004",
    pin: "456789",
    funcao: "Brigadista de Emergência",
    equipe: "Brigadistas de Emergência",
    turno: "1º Turno",
    unidade: "Unidade Industrial",
    status: "Ativo",
    presente: true,
    foto: "https://i.pravatar.cc/300?img=47",
  },
  {
    id: 5,
    nome: "Roberto Alves",
    matricula: "BRG005",
    pin: "567890",
    funcao: "Brigadista de Emergência",
    equipe: "Brigadistas de Emergência",
    turno: "1º Turno",
    unidade: "Unidade Industrial",
    status: "Ativo",
    presente: false,
    foto: "https://i.pravatar.cc/300?img=15",
  },
  {
    id: 6,
    nome: "Amanda Souza",
    matricula: "BRG006",
    pin: "678901",
    funcao: "Brigadista de Emergência",
    equipe: "Brigadistas de Emergência",
    turno: "1º Turno",
    unidade: "Unidade Industrial",
    status: "Ativo",
    presente: true,
    foto: "https://i.pravatar.cc/300?img=45",
  },
  {
    id: 7,
    nome: "Fernando Costa",
    matricula: "BRG007",
    pin: "789012",
    funcao: "Brigadista de Emergência",
    equipe: "Equipe de Abandono",
    turno: "1º Turno",
    unidade: "Unidade Industrial",
    status: "Ativo",
    presente: true,
    foto: "https://i.pravatar.cc/300?img=14",
  },
  {
    id: 8,
    nome: "Patrícia Rocha",
    matricula: "BRG008",
    pin: "890123",
    funcao: "Brigadista de Emergência",
    equipe: "Equipe de Abandono",
    turno: "1º Turno",
    unidade: "Unidade Industrial",
    status: "Ativo",
    presente: false,
    foto: "https://i.pravatar.cc/300?img=44",
  },
  {
    id: 9,
    nome: "Lucas Mendes",
    matricula: "BRG009",
    pin: "901234",
    funcao: "Brigadista de Emergência",
    equipe: "Equipe de Abandono",
    turno: "1º Turno",
    unidade: "Unidade Industrial",
    status: "Ativo",
    presente: true,
    foto: "https://i.pravatar.cc/300?img=33",
  },
  {
    id: 10,
    nome: "Ana Carolina",
    matricula: "BRG010",
    pin: "112233",
    funcao: "Brigadista de Emergência",
    equipe: "Primeiros Socorros",
    turno: "1º Turno",
    unidade: "Unidade Industrial",
    status: "Ativo",
    presente: true,
    foto: "https://i.pravatar.cc/300?img=32",
  },
  {
    id: 11,
    nome: "Rafael Martins",
    matricula: "BRG011",
    pin: "223344",
    funcao: "Brigadista de Emergência",
    equipe: "Primeiros Socorros",
    turno: "1º Turno",
    unidade: "Unidade Industrial",
    status: "Ativo",
    presente: false,
    foto: "https://i.pravatar.cc/300?img=52",
  },
  {
    id: 12,
    nome: "Camila Ferreira",
    matricula: "BRG012",
    pin: "334455",
    funcao: "Brigadista de Emergência",
    equipe: "Primeiros Socorros",
    turno: "1º Turno",
    unidade: "Unidade Industrial",
    status: "Ativo",
    presente: true,
    foto: "https://i.pravatar.cc/300?img=48",
  },
];

const STORAGE_KEY = "brigada-digital-dados";
const CONFIG_KEY = "brigada-digital-config";
const AUTH_KEY = "brigada-digital-admin-auth";

// LOGIN LOCAL TEMPORÁRIO.
// Para produção, substituiremos isto pelo Supabase Auth.
// Usuário: admin
// Senha: Brigada@2026
const ADMIN_USER = "admin";
const ADMIN_PASSWORD = "Brigada@2026";

function estaAutenticado() {
  return sessionStorage.getItem(AUTH_KEY) === "autenticado";
}

const configuracaoInicial = {
  empresaNome: "EMPRESA",
  unidadeNome: "Unidade Industrial",
  turno: "1º Turno",
  logoEmpresa: "",
  logoBrigada: "",
  latitude: "",
  longitude: "",
  raioGeofence: 150,
  precisaoMaxima: 120,
  avisoUrgenteAtivo: false,
  avisoUrgenteTexto: "",
  avisosEmergencia: [
    { id: 1, titulo: "Bombeiros", contato: "193", ativo: true },
    { id: 2, titulo: "SAMU", contato: "192", ativo: true },
    { id: 3, titulo: "Brigada / Ramal", contato: "2222", ativo: true },
  ],
  mensagemSeguranca: "Em caso de emergência, mantenha a calma e siga as orientações da Brigada.",
};

function carregarDados() {
  try {
    const salvo = localStorage.getItem(STORAGE_KEY);

    if (salvo) {
      const antigos = JSON.parse(salvo);

      // Migração automática: preserva presença/fotos já salvas e acrescenta
      // campos novos (como PIN) vindos dos dados iniciais.
      const baseMigrada = dadosIniciais.map((base) => {
        const antigo = antigos.find(
          (item) => item.id === base.id || item.matricula === base.matricula
        );
        return antigo ? { ...base, ...antigo, pin: antigo.pin || base.pin } : base;
      });

      // Preserva também pessoas criadas/importadas depois dos dados iniciais.
      const extras = antigos.filter(
        (antigo) =>
          !dadosIniciais.some(
            (base) =>
              base.id === antigo.id ||
              String(base.matricula) === String(antigo.matricula)
          )
      );

      return [...baseMigrada, ...extras];
    }
  } catch (erro) {
    console.error("Erro ao carregar dados:", erro);
  }

  return dadosIniciais;
}

function carregarConfiguracao() {
  try {
    const salvo = localStorage.getItem(CONFIG_KEY);
    if (salvo) return { ...configuracaoInicial, ...JSON.parse(salvo) };
  } catch (erro) {
    console.error("Erro ao carregar configurações:", erro);
  }
  return configuracaoInicial;
}

function PessoaCard({ pessoa, destaque = false }) {
  return (
    <div className={`pessoa-card ${pessoa.presente ? "presente" : "ausente"} ${destaque ? "destaque" : ""}`}>
      <div className="foto-container">
        {pessoa.foto ? (
          <img src={pessoa.foto} alt={pessoa.nome} />
        ) : (
          <div className="sem-foto"><Users size={34} /></div>
        )}
        <span className={`status-dot ${pessoa.presente ? "online" : "offline"}`} />
      </div>
      <div className="pessoa-info">
        <strong>{pessoa.nome}</strong>
        {pessoa.funcao && <span className="funcao">{pessoa.funcao}</span>}
        <span className={`status-text ${pessoa.presente ? "status-presente" : "status-ausente"}`}>
          {pessoa.presente ? "NA UNIDADE" : "AUSENTE"}
        </span>
      </div>
    </div>
  );
}

function Mural({ brigadistas, configuracao }) {
  const [agora, setAgora] = useState(new Date());

  useEffect(() => {
    const timer = setInterval(() => setAgora(new Date()), 1000);
    return () => clearInterval(timer);
  }, []);

  const ativos = brigadistas.filter(
    (pessoa) => String(pessoa.status).toLowerCase() !== "inativo"
  );
  const presentes = ativos.filter((pessoa) => pessoa.presente).length;
  const ausentes = ativos.length - presentes;

  const contatosEmergencia = (configuracao.avisosEmergencia || []).filter(
    (item) => item.ativo && (item.titulo || item.contato)
  );

  const coordenador = ativos.find((pessoa) =>
    pessoa.funcao.toLowerCase().includes("coordenador")
  );
  const lider = ativos.find(
    (pessoa) =>
      pessoa.funcao.toLowerCase().includes("líder") ||
      pessoa.funcao.toLowerCase().includes("lider")
  );

  const pessoasDasEquipes = ativos.filter(
    (pessoa) => pessoa.id !== coordenador?.id && pessoa.id !== lider?.id
  );
  const nomesEquipes = [
    ...new Set(pessoasDasEquipes.map((pessoa) => pessoa.equipe || "Brigadistas de Emergência")),
  ];

  const corEquipe = (nome) => {
    const n = nome.toLowerCase();
    if (n.includes("abandono")) return "laranja";
    if (n.includes("socorro")) return "verde";
    return "vermelho";
  };

  const iconeEquipe = (nome) => {
    const n = nome.toLowerCase();
    if (n.includes("abandono")) return "↗";
    if (n.includes("socorro")) return "✚";
    return "🔥";
  };

  return (
    <div className="mural-premium">
      <header className="premium-header">
        <div className="premium-brand">
          {configuracao.logoBrigada ? (
            <img src={configuracao.logoBrigada} alt="Logo da Brigada" />
          ) : (
            <ShieldCheck size={62} />
          )}
          <div className="lema-lateral">
            <span>PREVENIR</span><span>ORIENTAR</span><span>SOCORRER</span><span>SALVAR</span>
          </div>
        </div>

        <div className="premium-title">
          <div className="premium-kicker">SEGURANÇA • PREVENÇÃO • RESPOSTA</div>
          <h1>MURAL DIGITAL DA BRIGADA</h1>
          <div className="premium-sub">
            <b>BRIGADA DE EMERGÊNCIA</b><i />{configuracao.unidadeNome}<i />{configuracao.turno}
          </div>
        </div>

        <div className="premium-company">
          {configuracao.logoEmpresa ? (
            <img src={configuracao.logoEmpresa} alt={configuracao.empresaNome} />
          ) : (
            <strong>{configuracao.empresaNome}</strong>
          )}
        </div>
      </header>

      {configuracao.avisoUrgenteAtivo && configuracao.avisoUrgenteTexto && (
        <div className="mural-alerta-urgente">
          <AlertTriangle size={22} />
          <strong>AVISO URGENTE</strong>
          <span>{configuracao.avisoUrgenteTexto}</span>
        </div>
      )}

      <main className="premium-main">
        <section className="premium-lideranca">
          {coordenador && (
            <div className="premium-leader">
              <div className="leader-label red">★ <span>COORDENADOR<br />DE EMERGÊNCIA</span></div>
              <PessoaCard pessoa={coordenador} destaque />
            </div>
          )}
          <div className="leader-connector" />
          {lider && (
            <div className="premium-leader">
              <div className="leader-label blue">♢ <span>LÍDER DA BRIGADA<br />DE EMERGÊNCIA</span></div>
              <PessoaCard pessoa={lider} destaque />
            </div>
          )}
        </section>

        <section className="premium-equipes">
          {nomesEquipes.map((nomeEquipe) => {
            const pessoas = pessoasDasEquipes.filter((pessoa) => pessoa.equipe === nomeEquipe);
            const tema = corEquipe(nomeEquipe);
            return (
              <article className={`premium-equipe ${tema}`} key={nomeEquipe}>
                <div className="premium-equipe-title">
                  <span className="team-icon">{iconeEquipe(nomeEquipe)}</span>
                  <strong>{nomeEquipe}</strong>
                  <span className="team-count">{pessoas.length} MEMBROS</span>
                </div>
                <div className="premium-pessoas">
                  {pessoas.map((pessoa) => <PessoaCard key={pessoa.id} pessoa={pessoa} />)}
                </div>
              </article>
            );
          })}
        </section>
      </main>

      <footer className="premium-footer">
        <div className="premium-stats">
          <div className="stat green"><UserCheck size={20}/><span>Presentes</span><b>{presentes}</b></div>
          <div className="stat red"><UserX size={20}/><span>Ausentes</span><b>{ausentes}</b></div>
          <div className="stat blue"><Users size={20}/><span>Total</span><b>{ativos.length}</b></div>
        </div>

        <div className="premium-phrase">“BRIGADA FORTE, AMBIENTE SEGURO”</div>

        <div className="premium-clock">
          <Clock size={22}/>
          <div>
            <span>{agora.toLocaleDateString("pt-BR")}</span>
            <small>{agora.toLocaleDateString("pt-BR", { weekday: "long" })}</small>
          </div>
          <b>{agora.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" })}</b>
        </div>
      </footer>

      <div className="mural-emergency-bar">
        <div className="emergency-label">
          <Phone size={17} />
          <strong>EMERGÊNCIA</strong>
        </div>

        <div className="emergency-contacts">
          {contatosEmergencia.length > 0 ? (
            contatosEmergencia.map((item) => (
              <div className="emergency-contact" key={item.id}>
                <span>{item.titulo}</span>
                <b>{item.contato}</b>
              </div>
            ))
          ) : (
            <span className="emergency-empty">Contatos de emergência não cadastrados</span>
          )}
        </div>

        <div className="emergency-message">
          {configuracao.mensagemSeguranca || "Segurança em primeiro lugar."}
        </div>

        <span className="online-indicator"><i /> SISTEMA ONLINE</span>
      </div>
    </div>
  );
}

function Presenca({ brigadistas, registrarPresenca }) {
  const navigate = useNavigate();
  const [pin, setPin] = useState("");
  const [idSelecionado, setIdSelecionado] = useState(null);
  const [erro, setErro] = useState("");
  const [mensagem, setMensagem] = useState("");

  const pessoaSelecionada = brigadistas.find(
    (pessoa) => pessoa.id === idSelecionado
  );

  const abrirMural = () => {
    window.open("/mural", "muralBrigada", "width=1366,height=768");
  };

  const limparTerminal = () => {
    setPin("");
    setIdSelecionado(null);
    setErro("");
    setMensagem("");
  };

  const digitarNumero = (numero) => {
    if (pin.length >= 6 || mensagem) return;
    setErro("");
    setPin((anterior) => `${anterior}${numero}`.slice(0, 6));
  };

  const apagarNumero = () => {
    if (mensagem) return;
    setErro("");
    setPin((anterior) => anterior.slice(0, -1));
  };

  const validarPin = () => {
    if (pin.length !== 6) {
      setErro("Digite os 6 números do seu PIN.");
      return;
    }

    const pessoa = brigadistas.find(
      (item) =>
        String(item.status).toLowerCase() !== "inativo" &&
        String(item.pin || "") === pin
    );

    if (!pessoa) {
      setErro("PIN não localizado. Confira os números e tente novamente.");
      setPin("");
      return;
    }

    setErro("");
    setIdSelecionado(pessoa.id);
  };

  const registrar = (presente) => {
    if (!pessoaSelecionada) return;

    registrarPresenca(pessoaSelecionada.id, presente, "Painel/Terminal");

    const horario = new Date().toLocaleTimeString("pt-BR", {
      hour: "2-digit",
      minute: "2-digit",
    });

    setMensagem(
      presente
        ? `Entrada registrada com sucesso às ${horario}`
        : `Saída registrada com sucesso às ${horario}`
    );

    setTimeout(limparTerminal, 2200);
  };

  return (
    <div style={{
      minHeight: "100vh",
      background: "linear-gradient(135deg, #071b2b 0%, #0b2b3d 55%, #0b3445 100%)",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      padding: 20,
      fontFamily: "Arial, sans-serif",
    }}>
      <div style={{
        width: "100%",
        maxWidth: 520,
        background: "#ffffff",
        borderRadius: 22,
        overflow: "hidden",
        boxShadow: "0 24px 70px rgba(0,0,0,.35)",
      }}>
        <div style={{
          background: "#071b2b",
          color: "#fff",
          padding: "22px 26px",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
        }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <ShieldCheck size={38} />
            <div>
              <div style={{ fontSize: 12, letterSpacing: 2, opacity: 0.75 }}>
                BRIGADA DIGITAL
              </div>
              <strong style={{ fontSize: 20 }}>REGISTRO DE PRESENÇA</strong>
            </div>
          </div>

          <button type="button" onClick={abrirMural} title="Abrir Gestão à Vista"
            style={{
              width: 44, height: 44, borderRadius: 10,
              border: "1px solid rgba(255,255,255,.25)",
              background: "rgba(255,255,255,.08)", color: "#fff", cursor: "pointer",
            }}>
            <Monitor size={22} />
          </button>
        </div>

        <div style={{ padding: 30 }}>
          {!pessoaSelecionada && !mensagem && (
            <>
              <div style={{ textAlign: "center", marginBottom: 22 }}>
                <div style={{
                  width: 68, height: 68, margin: "0 auto 14px", borderRadius: "50%",
                  background: "#edf4f6", display: "flex", alignItems: "center",
                  justifyContent: "center", color: "#0b3445",
                }}>
                  <LogIn size={32} />
                </div>
                <h2 style={{ margin: 0, color: "#14232c", fontSize: 25 }}>
                  Identifique-se
                </h2>
                <p style={{ margin: "8px 0 0", color: "#687780" }}>
                  Digite seu PIN de 6 dígitos
                </p>
              </div>

              <div style={{
                display: "flex", justifyContent: "center", gap: 13, margin: "22px 0 24px",
              }}>
                {[0, 1, 2, 3, 4, 5].map((indice) => (
                  <div key={indice} style={{
                    width: 20, height: 20, borderRadius: "50%",
                    background: indice < pin.length ? "#0b3445" : "#d9e1e5",
                    boxShadow: indice < pin.length
                      ? "0 0 0 4px rgba(11,52,69,.10)" : "none",
                  }} />
                ))}
              </div>

              {erro && (
                <div style={{
                  background: "#fff0f0", border: "1px solid #efc3c3",
                  color: "#a8242c", borderRadius: 10, padding: 12,
                  textAlign: "center", fontWeight: 700, marginBottom: 16,
                }}>
                  {erro}
                </div>
              )}

              <div style={{
                maxWidth: 330, margin: "0 auto", display: "grid",
                gridTemplateColumns: "repeat(3, 1fr)", gap: 11,
              }}>
                {[1,2,3,4,5,6,7,8,9].map((numero) => (
                  <button key={numero} type="button" onClick={() => digitarNumero(numero)}
                    style={{
                      height: 66, border: "1px solid #d4dde2", borderRadius: 13,
                      background: "#f8fafb", color: "#14232c", fontSize: 27,
                      fontWeight: 800, cursor: "pointer",
                      boxShadow: "0 3px 8px rgba(0,0,0,.05)",
                    }}>
                    {numero}
                  </button>
                ))}

                <button type="button" onClick={() => { setPin(""); setErro(""); }}
                  style={{
                    height: 66, border: "1px solid #e0c7c7", borderRadius: 13,
                    background: "#fff5f5", color: "#a8242c", fontSize: 14,
                    fontWeight: 900, cursor: "pointer",
                  }}>
                  LIMPAR
                </button>

                <button type="button" onClick={() => digitarNumero(0)}
                  style={{
                    height: 66, border: "1px solid #d4dde2", borderRadius: 13,
                    background: "#f8fafb", color: "#14232c", fontSize: 27,
                    fontWeight: 800, cursor: "pointer",
                  }}>
                  0
                </button>

                <button type="button" onClick={apagarNumero}
                  style={{
                    height: 66, border: "1px solid #d4dde2", borderRadius: 13,
                    background: "#f1f4f6", color: "#14232c", fontSize: 25,
                    fontWeight: 900, cursor: "pointer",
                  }}>
                  ⌫
                </button>
              </div>

              <button type="button" onClick={validarPin} disabled={pin.length !== 6}
                style={{
                  width: "100%", height: 56, marginTop: 20, border: 0,
                  borderRadius: 12, background: pin.length === 6 ? "#d9272e" : "#cbd3d7",
                  color: "#fff", fontSize: 16, fontWeight: 900,
                  cursor: pin.length === 6 ? "pointer" : "not-allowed",
                  letterSpacing: 0.7,
                }}>
                CONFIRMAR
              </button>

              <div style={{
                textAlign: "center", color: "#95a0a6", fontSize: 11, marginTop: 16,
              }}>
                Seu PIN não será exibido na tela.
              </div>
            </>
          )}

          {pessoaSelecionada && !mensagem && (
            <div style={{ textAlign: "center" }}>
              <div style={{
                width: 142, height: 142, margin: "0 auto 16px",
                borderRadius: "50%", overflow: "hidden", background: "#e5e9ec",
                border: pessoaSelecionada.presente
                  ? "5px solid #20a45b" : "5px solid #c93038",
                display: "flex", alignItems: "center", justifyContent: "center",
              }}>
                {pessoaSelecionada.foto ? (
                  <img src={pessoaSelecionada.foto} alt={pessoaSelecionada.nome}
                    style={{
                      width: "100%", height: "100%", objectFit: "cover",
                      filter: pessoaSelecionada.presente ? "none" : "grayscale(100%)",
                    }} />
                ) : (
                  <Users size={58} color="#7b858d" />
                )}
              </div>

              <h2 style={{ margin: "0 0 5px", color: "#14232c", fontSize: 27 }}>
                {pessoaSelecionada.nome}
              </h2>
              <div style={{ color: "#687780", fontWeight: 700, marginBottom: 4 }}>
                {pessoaSelecionada.matricula}
              </div>
              <div style={{ color: "#687780", marginBottom: 3 }}>
                {pessoaSelecionada.funcao}
              </div>
              <div style={{ color: "#687780", marginBottom: 20 }}>
                {pessoaSelecionada.equipe}
              </div>

              <div style={{
                display: "inline-flex", alignItems: "center", gap: 8,
                padding: "9px 15px", borderRadius: 30, marginBottom: 22,
                background: pessoaSelecionada.presente ? "#e4f6eb" : "#fde9ea",
                color: pessoaSelecionada.presente ? "#168143" : "#a8242c",
                fontWeight: 900, fontSize: 13,
              }}>
                ● {pessoaSelecionada.presente ? "NA UNIDADE" : "AUSENTE"}
              </div>

              {!pessoaSelecionada.presente ? (
                <button type="button" onClick={() => registrar(true)}
                  style={{
                    width: "100%", height: 62, border: 0, borderRadius: 12,
                    background: "#168143", color: "#fff", fontSize: 17,
                    fontWeight: 900, cursor: "pointer",
                  }}>
                  <LogIn size={21} style={{ verticalAlign: "middle", marginRight: 8 }} />
                  ESTOU NA UNIDADE
                </button>
              ) : (
                <button type="button" onClick={() => registrar(false)}
                  style={{
                    width: "100%", height: 62, border: 0, borderRadius: 12,
                    background: "#b71f29", color: "#fff", fontSize: 17,
                    fontWeight: 900, cursor: "pointer",
                  }}>
                  <LogOut size={21} style={{ verticalAlign: "middle", marginRight: 8 }} />
                  REGISTRAR MINHA SAÍDA
                </button>
              )}

              <button type="button" onClick={limparTerminal}
                style={{
                  width: "100%", height: 46, marginTop: 11,
                  border: "1px solid #ccd3d8", borderRadius: 10,
                  background: "#fff", color: "#43515a", fontWeight: 800, cursor: "pointer",
                }}>
                CANCELAR
              </button>
            </div>
          )}

          {mensagem && (
            <div style={{
              minHeight: 430, display: "flex", flexDirection: "column",
              alignItems: "center", justifyContent: "center", textAlign: "center",
            }}>
              <div style={{
                width: 88, height: 88, borderRadius: "50%", background: "#e2f6ea",
                color: "#168143", display: "flex", alignItems: "center",
                justifyContent: "center", fontSize: 46, fontWeight: 900, marginBottom: 20,
              }}>
                ✓
              </div>
              <h2 style={{ color: "#14232c", margin: "0 0 8px" }}>
                REGISTRO REALIZADO
              </h2>
              <strong style={{ fontSize: 20, color: "#263740" }}>
                {pessoaSelecionada?.nome}
              </strong>
              <p style={{ color: "#168143", fontWeight: 800, marginTop: 12 }}>
                {mensagem}
              </p>
              <span style={{ color: "#8a969c", fontSize: 12 }}>
                Retornando ao terminal...
              </span>
            </div>
          )}
        </div>

        <div style={{
          borderTop: "1px solid #e7ecef", padding: "12px 18px",
          display: "flex", justifyContent: "space-between",
          alignItems: "center", color: "#87949b", fontSize: 11,
        }}>
          <span>Unidade Industrial</span>
          <button type="button" onClick={() => navigate("/admin")}
            style={{
              border: 0, background: "transparent", color: "#aeb8bd",
              fontSize: 10, cursor: "pointer",
            }}>
            ADMINISTRAÇÃO
          </button>
        </div>
      </div>
    </div>
  );
}



function calcularDistanciaMetros(lat1, lon1, lat2, lon2) {
  const R = 6371000;
  const rad = (graus) => (graus * Math.PI) / 180;
  const dLat = rad(lat2 - lat1);
  const dLon = rad(lon2 - lon1);
  const a =
    Math.sin(dLat / 2) ** 2 +
    Math.cos(rad(lat1)) *
      Math.cos(rad(lat2)) *
      Math.sin(dLon / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function Geolocalizacao({ brigadistas, configuracao, registrarPresenca }) {
  const navigate = useNavigate();
  const [pin, setPin] = useState("");
  const [brigadistaId, setBrigadistaId] = useState(null);
  const [erro, setErro] = useState("");
  const [statusGps, setStatusGps] = useState("aguardando");
  const [distancia, setDistancia] = useState(null);
  const [precisao, setPrecisao] = useState(null);
  const [mensagem, setMensagem] = useState(
    "Identifique-se para ativar a presença por localização."
  );

  const brigadista = brigadistas.find((p) => p.id === brigadistaId);
  const latitudeEmpresa = Number(configuracao.latitude);
  const longitudeEmpresa = Number(configuracao.longitude);
  const raio = Number(configuracao.raioGeofence || 150);
  const precisaoMaxima = Number(configuracao.precisaoMaxima || 120);
  const localConfigurado =
    Number.isFinite(latitudeEmpresa) &&
    Number.isFinite(longitudeEmpresa) &&
    configuracao.latitude !== "" &&
    configuracao.longitude !== "";

  const validarPinGeo = () => {
    const pessoa = brigadistas.find(
      (p) =>
        String(p.status).toLowerCase() !== "inativo" &&
        String(p.pin || "") === pin
    );

    if (!pessoa) {
      setErro("PIN não localizado.");
      setPin("");
      return;
    }

    setErro("");
    setBrigadistaId(pessoa.id);
    setMensagem("Ative a localização para iniciar a detecção.");
  };

  useEffect(() => {
    if (!brigadistaId) return;

    if (!localConfigurado) {
      setStatusGps("erro");
      setMensagem("A localização da empresa ainda não foi configurada pelo administrador.");
      return;
    }

    if (!navigator.geolocation) {
      setStatusGps("erro");
      setMensagem("Este aparelho/navegador não oferece geolocalização.");
      return;
    }

    setStatusGps("localizando");
    setMensagem("Localizando o aparelho...");

    const watchId = navigator.geolocation.watchPosition(
      (posicao) => {
        const { latitude, longitude, accuracy } = posicao.coords;
        const d = calcularDistanciaMetros(
          latitude,
          longitude,
          latitudeEmpresa,
          longitudeEmpresa
        );

        setDistancia(Math.round(d));
        setPrecisao(Math.round(accuracy));

        if (accuracy > precisaoMaxima) {
          setStatusGps("impreciso");
          setMensagem(
            `Sinal de localização impreciso (${Math.round(accuracy)} m). Nenhuma presença foi alterada.`
          );
          return;
        }

        const pessoaAtual = brigadistas.find((p) => p.id === brigadistaId);
        if (!pessoaAtual) return;

        // Histerese: entra no raio configurado, mas só registra saída
        // quando ultrapassar o raio + 30 m. Evita alternância por oscilação do GPS.
        const dentro = d <= raio;
        const foraComMargem = d > raio + 30;

        if (dentro) {
          setStatusGps("dentro");
          setMensagem(`Você está dentro da área da ${configuracao.unidadeNome}.`);
          if (!pessoaAtual.presente) {
            registrarPresenca(brigadistaId, true, "Geolocalização");
          }
        } else if (foraComMargem) {
          setStatusGps("fora");
          setMensagem(`Você está fora da área da ${configuracao.unidadeNome}.`);
          if (pessoaAtual.presente && pessoaAtual.origemPresenca === "Geolocalização") {
            registrarPresenca(brigadistaId, false, "Geolocalização");
          }
        } else {
          setStatusGps("limite");
          setMensagem("Você está próximo ao limite da área. Aguardando confirmação do GPS.");
        }
      },
      (geoErro) => {
        setStatusGps("erro");
        if (geoErro.code === 1) {
          setMensagem("Permissão de localização negada. Use o painel manual da recepção.");
        } else {
          setMensagem("Não foi possível obter a localização. Use o painel manual da recepção.");
        }
      },
      {
        enableHighAccuracy: true,
        maximumAge: 15000,
        timeout: 20000,
      }
    );

    return () => navigator.geolocation.clearWatch(watchId);
  }, [
    brigadistaId,
    localConfigurado,
    latitudeEmpresa,
    longitudeEmpresa,
    raio,
    precisaoMaxima,
    configuracao.unidadeNome,
    brigadistas,
    registrarPresenca,
  ]);

  if (!brigadistaId) {
    return (
      <div className="geo-page">
        <div className="geo-card">
          <div className="geo-icon"><Smartphone size={38} /></div>
          <h1>Presença Inteligente</h1>
          <p>
            Informe seu PIN de 6 dígitos. Depois, permita o acesso à localização
            do aparelho.
          </p>

          <input
            className="geo-pin"
            type="password"
            inputMode="numeric"
            maxLength={6}
            value={pin}
            onChange={(e) => {
              setPin(e.target.value.replace(/\D/g, "").slice(0, 6));
              setErro("");
            }}
            placeholder="••••••"
          />

          {erro && <div className="geo-error">{erro}</div>}

          <button
            className="geo-primary"
            disabled={pin.length !== 6}
            onClick={validarPinGeo}
          >
            <Navigation size={19} /> ATIVAR MINHA LOCALIZAÇÃO
          </button>

          <button className="geo-manual" onClick={() => navigate("/presenca")}>
            ESTOU SEM CELULAR / USAR PAINEL MANUAL
          </button>

          <div className="geo-privacy">
            <ShieldCheck size={16} />
            A Brigada Digital usa a localização apenas para verificar se o
            aparelho está dentro da área cadastrada da unidade.
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="geo-page">
      <div className="geo-card">
        <div className={`geo-icon ${statusGps}`}>
          <MapPin size={40} />
        </div>

        {brigadista?.foto && (
          <img className="geo-photo" src={brigadista.foto} alt={brigadista.nome} />
        )}

        <h1>{brigadista?.nome}</h1>
        <p>{brigadista?.funcao}</p>

        <div className={`geo-status ${statusGps}`}>
          {statusGps === "dentro" && "● NA UNIDADE"}
          {statusGps === "fora" && "● FORA DA UNIDADE"}
          {statusGps === "limite" && "● PRÓXIMO AO LIMITE"}
          {statusGps === "impreciso" && "● LOCALIZAÇÃO IMPRECISA"}
          {statusGps === "localizando" && "● LOCALIZANDO..."}
          {statusGps === "erro" && "● LOCALIZAÇÃO INDISPONÍVEL"}
          {statusGps === "aguardando" && "● AGUARDANDO"}
        </div>

        <div className="geo-message">{mensagem}</div>

        <div className="geo-metrics">
          <div><span>Distância da unidade</span><b>{distancia === null ? "—" : `${distancia} m`}</b></div>
          <div><span>Precisão do GPS</span><b>{precisao === null ? "—" : `${precisao} m`}</b></div>
          <div><span>Raio configurado</span><b>{raio} m</b></div>
        </div>

        <button
          className="geo-manual"
          onClick={() => navigate("/presenca")}
        >
          USAR PAINEL MANUAL
        </button>

        <button
          className="geo-exit"
          onClick={() => {
            setBrigadistaId(null);
            setPin("");
            setDistancia(null);
            setPrecisao(null);
            setStatusGps("aguardando");
          }}
        >
          ENCERRAR NESTE APARELHO
        </button>
      </div>
    </div>
  );
}

function Login({ onLogin }) {
  const navigate = useNavigate();
  const [usuario, setUsuario] = useState("");
  const [senha, setSenha] = useState("");
  const [mostrarSenha, setMostrarSenha] = useState(false);
  const [erro, setErro] = useState("");

  const entrar = (e) => {
    e.preventDefault();

    if (usuario.trim() === ADMIN_USER && senha === ADMIN_PASSWORD) {
      sessionStorage.setItem(AUTH_KEY, "autenticado");
      setErro("");
      onLogin(true);
      navigate("/admin", { replace: true });
      return;
    }

    setErro("Usuário ou senha incorretos.");
    setSenha("");
  };

  return (
    <div className="login-page">
      <div className="login-card">
        <div className="login-logo">
          <div className="login-shield"><ShieldCheck size={42} /></div>
          <div>
            <span>BRIGADA</span>
            <strong>DIGITAL</strong>
          </div>
        </div>

        <div className="login-heading">
          <div className="login-lock"><LockKeyhole size={25} /></div>
          <h1>Acesso Administrativo</h1>
          <p>Entre com suas credenciais para acessar a gestão da Brigada Digital.</p>
        </div>

        <form onSubmit={entrar}>
          <label className="login-label">
            Usuário
            <input
              className="login-input"
              value={usuario}
              onChange={(e) => {
                setUsuario(e.target.value);
                setErro("");
              }}
              placeholder="Digite seu usuário"
              autoComplete="username"
              autoFocus
            />
          </label>

          <label className="login-label">
            Senha
            <div className="login-password-wrap">
              <input
                className="login-input"
                type={mostrarSenha ? "text" : "password"}
                value={senha}
                onChange={(e) => {
                  setSenha(e.target.value);
                  setErro("");
                }}
                placeholder="Digite sua senha"
                autoComplete="current-password"
              />
              <button
                type="button"
                className="login-eye"
                onClick={() => setMostrarSenha((v) => !v)}
                aria-label={mostrarSenha ? "Ocultar senha" : "Mostrar senha"}
              >
                {mostrarSenha ? <EyeOff size={20} /> : <Eye size={20} />}
              </button>
            </div>
          </label>

          {erro && <div className="login-error">{erro}</div>}

          <button className="login-submit" type="submit">
            <LockKeyhole size={18} />
            ENTRAR NO SISTEMA
          </button>
        </form>

        <button className="login-back" type="button" onClick={() => navigate("/presenca")}>
          Voltar ao registro de presença
        </button>

        <div className="login-security">
          <ShieldCheck size={16} />
          Sessão administrativa protegida neste navegador
        </div>
      </div>
    </div>
  );
}

function AdminProtegido({ autenticado, children }) {
  if (!autenticado) {
    return <Navigate to="/login" replace />;
  }
  return children;
}

function Admin({
  configuracao,
  salvarConfiguracao,
  brigadistas,
  atualizarBrigadistas,
  onLogout,
}) {
  const navigate = useNavigate();
  const [aba, setAba] = useState("brigadistas");
  const [form, setForm] = useState(configuracao);
  const [salvo, setSalvo] = useState(false);
  const [busca, setBusca] = useState("");
  const [mensagemImportacao, setMensagemImportacao] = useState("");

  useEffect(() => {
    setForm(configuracao);
  }, [configuracao]);

  const abrirMural = () => {
    window.open("/mural", "muralBrigada", "width=1366,height=768");
  };

  const lerImagemConfig = (arquivo, campo) => {
    if (!arquivo) return;
    const reader = new FileReader();
    reader.onload = () =>
      setForm((anterior) => ({ ...anterior, [campo]: reader.result }));
    reader.readAsDataURL(arquivo);
  };

  const salvarConfig = () => {
    salvarConfiguracao(form);
    setSalvo(true);
    setTimeout(() => setSalvo(false), 1800);
  };

  const capturarLocalEmpresa = () => {
    if (!navigator.geolocation) {
      window.alert("Este navegador não oferece geolocalização.");
      return;
    }

    navigator.geolocation.getCurrentPosition(
      (posicao) => {
        setForm((anterior) => ({
          ...anterior,
          latitude: posicao.coords.latitude.toFixed(7),
          longitude: posicao.coords.longitude.toFixed(7),
        }));
      },
      () => {
        window.alert(
          "Não foi possível obter a localização. Autorize a localização no navegador ou preencha latitude e longitude manualmente."
        );
      },
      { enableHighAccuracy: true, timeout: 20000, maximumAge: 0 }
    );
  };

  const adicionarContatoEmergencia = () => {
    const contatos = form.avisosEmergencia || [];
    const novoId =
      contatos.length > 0
        ? Math.max(...contatos.map((item) => Number(item.id) || 0)) + 1
        : 1;

    setForm((anterior) => ({
      ...anterior,
      avisosEmergencia: [
        ...(anterior.avisosEmergencia || []),
        { id: novoId, titulo: "Novo contato", contato: "", ativo: true },
      ],
    }));
  };

  const atualizarContatoEmergencia = (id, campo, valor) => {
    setForm((anterior) => ({
      ...anterior,
      avisosEmergencia: (anterior.avisosEmergencia || []).map((item) =>
        item.id === id ? { ...item, [campo]: valor } : item
      ),
    }));
  };

  const removerContatoEmergencia = (id) => {
    setForm((anterior) => ({
      ...anterior,
      avisosEmergencia: (anterior.avisosEmergencia || []).filter(
        (item) => item.id !== id
      ),
    }));
  };

  const atualizarPessoa = (id, campo, valor) => {
    atualizarBrigadistas(
      brigadistas.map((pessoa) =>
        pessoa.id === id ? { ...pessoa, [campo]: valor } : pessoa
      )
    );
  };

  const carregarFoto = (id, arquivo) => {
    if (!arquivo) return;
    const reader = new FileReader();
    reader.onload = () => atualizarPessoa(id, "foto", reader.result);
    reader.readAsDataURL(arquivo);
  };

  const removerPessoa = (id) => {
    if (!window.confirm("Deseja realmente remover este brigadista?")) return;
    atualizarBrigadistas(brigadistas.filter((pessoa) => pessoa.id !== id));
  };

  const novoBrigadista = () => {
    const proximoId =
      brigadistas.length > 0
        ? Math.max(...brigadistas.map((p) => Number(p.id) || 0)) + 1
        : 1;

    atualizarBrigadistas([
      ...brigadistas,
      {
        id: proximoId,
        nome: "Novo Brigadista",
        matricula: "",
        pin: "",
        funcao: "Brigadista de Emergência",
        equipe: "Brigadistas de Emergência",
        turno: configuracao.turno || "1º Turno",
        unidade: configuracao.unidadeNome || "Unidade Industrial",
        status: "Ativo",
        presente: false,
        foto: "",
      },
    ]);
  };

  const importarExcel = (arquivo) => {
    if (!arquivo) return;

    const reader = new FileReader();
    reader.onload = (evento) => {
      try {
        const workbook = XLSX.read(evento.target.result, { type: "array" });
        const primeiraAba = workbook.Sheets[workbook.SheetNames[0]];
        const linhas = XLSX.utils.sheet_to_json(primeiraAba, { defval: "" });

        if (!linhas.length) {
          setMensagemImportacao("A planilha está vazia.");
          return;
        }

        const matriculasExistentes = new Set(
          brigadistas.map((p) => String(p.matricula).trim().toLowerCase())
        );

        let proximoId =
          brigadistas.length > 0
            ? Math.max(...brigadistas.map((p) => Number(p.id) || 0)) + 1
            : 1;

        const novos = [];
        const ignorados = [];

        linhas.forEach((linha, indice) => {
          const nome = String(linha["Nome"] || "").trim();
          const matricula = String(
            linha["Matrícula"] || linha["Matricula"] || ""
          ).trim();
          const pin = String(linha["PIN"] || "").replace(/\D/g, "").slice(0, 6);
          const funcao = String(
            linha["Função"] || linha["Funcao"] || "Brigadista de Emergência"
          ).trim();
          const equipe = String(
            linha["Equipe"] || "Brigadistas de Emergência"
          ).trim();
          const turno = String(
            linha["Turno"] || configuracao.turno || "1º Turno"
          ).trim();
          const unidade = String(
            linha["Unidade"] || configuracao.unidadeNome || "Unidade Industrial"
          ).trim();
          const status = String(linha["Status"] || "Ativo").trim();

          if (!nome || !matricula) {
            ignorados.push(`linha ${indice + 2}: nome ou matrícula ausente`);
            return;
          }

          const chave = matricula.toLowerCase();
          if (matriculasExistentes.has(chave)) {
            ignorados.push(`linha ${indice + 2}: matrícula ${matricula} duplicada`);
            return;
          }

          matriculasExistentes.add(chave);
          novos.push({
            id: proximoId++,
            nome,
            matricula,
            pin,
            funcao,
            equipe,
            turno,
            unidade,
            status,
            presente: false,
            foto: "",
          });
        });

        if (novos.length) {
          atualizarBrigadistas([...brigadistas, ...novos]);
        }

        setMensagemImportacao(
          `${novos.length} brigadista(s) importado(s).` +
            (ignorados.length ? ` ${ignorados.length} linha(s) ignorada(s).` : "")
        );
      } catch (erro) {
        console.error(erro);
        setMensagemImportacao("Não foi possível ler a planilha. Use o modelo .xlsx.");
      }
    };
    reader.readAsArrayBuffer(arquivo);
  };

  const filtrados = brigadistas.filter((pessoa) => {
    const termo = busca.toLowerCase();
    return (
      pessoa.nome.toLowerCase().includes(termo) ||
      String(pessoa.matricula).toLowerCase().includes(termo) ||
      String(pessoa.funcao).toLowerCase().includes(termo) ||
      String(pessoa.equipe).toLowerCase().includes(termo)
    );
  });

  const campo = {
    width: "100%",
    height: 42,
    border: "1px solid #ccd3d8",
    borderRadius: 8,
    padding: "0 10px",
    boxSizing: "border-box",
    background: "#fff",
  };

  return (
    <div className="admin-app">
      <aside className="sidebar">
        <div className="sidebar-marca">
          <ShieldCheck size={35} />
          <div><strong>BRIGADA</strong><span>DIGITAL</span></div>
        </div>

        <button onClick={() => navigate("/presenca")}>
          <UserCheck size={19} /> Presença
        </button>
        <button onClick={() => navigate("/geolocalizacao")}>
          <MapPin size={19} /> Geolocalização
        </button>
        <button
          className={aba === "brigadistas" ? "ativo" : ""}
          onClick={() => setAba("brigadistas")}
        >
          <Users size={19} /> Brigadistas
        </button>
        <button
          className={aba === "avisos" ? "ativo" : ""}
          onClick={() => setAba("avisos")}
        >
          <Bell size={19} /> Avisos e Emergência
        </button>
        <button
          className={aba === "configuracoes" ? "ativo" : ""}
          onClick={() => setAba("configuracoes")}
        >
          <Settings size={19} /> Configurações
        </button>
        <button onClick={abrirMural}>
          <Monitor size={19} /> Gestão à Vista
        </button>

        <button
          className="sidebar-sair"
          onClick={() => {
            sessionStorage.removeItem(AUTH_KEY);
            onLogout();
            navigate("/login", { replace: true });
          }}
        >
          <LogOut size={19} /> Sair
        </button>
      </aside>

      <div className="admin-conteudo">
        <div className="admin-topo">
          <div>
            <span>PAINEL ADMINISTRATIVO</span>
            <h1>
              {aba === "brigadistas"
                ? "Gestão dos Brigadistas"
                : aba === "avisos"
                  ? "Avisos e Emergência"
                  : "Configurações do Mural"}
            </h1>
          </div>
          <button className="voltar-mural" onClick={abrirMural}>
            <ExternalLink size={18} /> ABRIR GESTÃO À VISTA
          </button>
        </div>

        {aba === "brigadistas" && (
          <>
            <section className="painel-card admin-toolbar-card">
              <div className="brigada-toolbar">
                <div className="brigada-search">
                  <Search size={18} />
                  <input
                    value={busca}
                    onChange={(e) => setBusca(e.target.value)}
                    placeholder="Buscar por nome, matrícula, função ou equipe..."
                  />
                </div>

                <label className="admin-action secondary">
                  <Upload size={18} />
                  IMPORTAR EXCEL
                  <input
                    type="file"
                    accept=".xlsx,.xls"
                    hidden
                    onChange={(e) => {
                      importarExcel(e.target.files?.[0]);
                      e.target.value = "";
                    }}
                  />
                </label>

                <button className="admin-action primary" onClick={novoBrigadista}>
                  <Plus size={18} /> NOVO BRIGADISTA
                </button>
              </div>

              {mensagemImportacao && (
                <div className="import-message">{mensagemImportacao}</div>
              )}

              <div className="import-hint">
                <FileSpreadsheet size={18} />
                <span>
                  Importação: Nome, Matrícula, PIN, Função, Equipe, Turno, Unidade e Status.
                  As fotos são adicionadas depois, individualmente.
                </span>
              </div>
            </section>

            <section className="painel-card">
              <div className="painel-titulo">
                <Users size={22} />
                <div>
                  <h2>Brigadistas cadastrados</h2>
                  <p>{brigadistas.length} cadastro(s) • {filtrados.length} exibido(s)</p>
                </div>
              </div>

              <div className="brigadistas-admin-grid">
                {filtrados.map((pessoa) => (
                  <article className="brigadista-admin-card" key={pessoa.id}>
                    <div className="admin-photo-box">
                      {pessoa.foto ? (
                        <img src={pessoa.foto} alt={pessoa.nome} />
                      ) : (
                        <div className="admin-no-photo"><Users size={42} /></div>
                      )}

                      <label className="photo-button" title="Alterar foto">
                        <Camera size={17} />
                        FOTO
                        <input
                          type="file"
                          accept="image/*"
                          hidden
                          onChange={(e) => {
                            carregarFoto(pessoa.id, e.target.files?.[0]);
                            e.target.value = "";
                          }}
                        />
                      </label>
                    </div>

                    <div className="admin-person-fields">
                      <label>
                        Nome
                        <input
                          value={pessoa.nome}
                          onChange={(e) => atualizarPessoa(pessoa.id, "nome", e.target.value)}
                        />
                      </label>

                      <div className="two-fields">
                        <label>
                          Matrícula
                          <input
                            value={pessoa.matricula}
                            onChange={(e) => atualizarPessoa(pessoa.id, "matricula", e.target.value)}
                          />
                        </label>
                        <label>
                          PIN
                          <input
                            inputMode="numeric"
                            maxLength={6}
                            value={pessoa.pin || ""}
                            onChange={(e) =>
                              atualizarPessoa(
                                pessoa.id,
                                "pin",
                                e.target.value.replace(/\D/g, "").slice(0, 6)
                              )
                            }
                          />
                        </label>
                      </div>

                      <label>
                        Função
                        <input
                          value={pessoa.funcao}
                          onChange={(e) => atualizarPessoa(pessoa.id, "funcao", e.target.value)}
                          placeholder="Ex.: Brigadista de Emergência"
                        />
                      </label>

                      <label>
                        Equipe
                        <select
                          value={pessoa.equipe}
                          onChange={(e) => atualizarPessoa(pessoa.id, "equipe", e.target.value)}
                        >
                          <option>Coordenação</option>
                          <option>Liderança</option>
                          <option>Brigadistas de Emergência</option>
                          <option>Equipe de Abandono</option>
                          <option>Primeiros Socorros</option>
                          <option>Combate</option>
                          <option>Isolamento</option>
                          <option>Comunicação</option>
                        </select>
                      </label>

                      <div className="two-fields">
                        <label>
                          Turno
                          <input
                            value={pessoa.turno}
                            onChange={(e) => atualizarPessoa(pessoa.id, "turno", e.target.value)}
                          />
                        </label>
                        <label>
                          Status
                          <select
                            value={pessoa.status}
                            onChange={(e) => atualizarPessoa(pessoa.id, "status", e.target.value)}
                          >
                            <option>Ativo</option>
                            <option>Inativo</option>
                          </select>
                        </label>
                      </div>

                      <label>
                        Unidade
                        <input
                          value={pessoa.unidade}
                          onChange={(e) => atualizarPessoa(pessoa.id, "unidade", e.target.value)}
                        />
                      </label>

                      <div className="admin-card-actions">
                        <span className="autosave-label"><Save size={15}/> SALVO AUTOMATICAMENTE</span>
                        <button className="delete-person" onClick={() => removerPessoa(pessoa.id)}>
                          <Trash2 size={16}/> REMOVER
                        </button>
                      </div>
                    </div>
                  </article>
                ))}
              </div>
            </section>
          </>
        )}

        {aba === "avisos" && (
          <>
            <section className="painel-card">
              <div className="painel-titulo">
                <Phone size={22} />
                <div>
                  <h2>Telefones e ramais de emergência</h2>
                  <p>
                    Cadastre os contatos que ficarão visíveis permanentemente no mural.
                  </p>
                </div>
              </div>

              <div className="emergency-admin-list">
                {(form.avisosEmergencia || []).map((item) => (
                  <div className="emergency-admin-row" key={item.id}>
                    <label>
                      Identificação
                      <input
                        value={item.titulo}
                        onChange={(e) =>
                          atualizarContatoEmergencia(item.id, "titulo", e.target.value)
                        }
                        placeholder="Ex.: Ambulatório"
                      />
                    </label>

                    <label>
                      Telefone / Ramal
                      <input
                        value={item.contato}
                        onChange={(e) =>
                          atualizarContatoEmergencia(item.id, "contato", e.target.value)
                        }
                        placeholder="Ex.: 2222"
                      />
                    </label>

                    <label className="emergency-active">
                      <span>Exibir no mural</span>
                      <input
                        type="checkbox"
                        checked={Boolean(item.ativo)}
                        onChange={(e) =>
                          atualizarContatoEmergencia(item.id, "ativo", e.target.checked)
                        }
                      />
                    </label>

                    <button
                      type="button"
                      className="emergency-delete"
                      onClick={() => removerContatoEmergencia(item.id)}
                      title="Remover contato"
                    >
                      <Trash2 size={18} />
                    </button>
                  </div>
                ))}
              </div>

              <button
                type="button"
                className="admin-action primary emergency-add"
                onClick={adicionarContatoEmergencia}
              >
                <Plus size={18} /> ADICIONAR TELEFONE / RAMAL
              </button>
            </section>

            <section className="painel-card emergency-message-card">
              <div className="painel-titulo">
                <ShieldCheck size={22} />
                <div>
                  <h2>Mensagem de segurança</h2>
                  <p>Texto curto exibido junto aos contatos de emergência.</p>
                </div>
              </div>

              <textarea
                className="emergency-textarea"
                maxLength={180}
                value={form.mensagemSeguranca || ""}
                onChange={(e) =>
                  setForm({ ...form, mensagemSeguranca: e.target.value })
                }
                placeholder="Em caso de emergência, mantenha a calma..."
              />
              <small className="char-counter">
                {(form.mensagemSeguranca || "").length}/180
              </small>
            </section>

            <section className={`painel-card urgent-admin-card ${form.avisoUrgenteAtivo ? "active" : ""}`}>
              <div className="urgent-admin-header">
                <div className="painel-titulo">
                  <AlertTriangle size={23} />
                  <div>
                    <h2>Aviso urgente</h2>
                    <p>
                      Quando ativado, aparece em destaque no topo do mural.
                    </p>
                  </div>
                </div>

                <label className="urgent-switch">
                  <input
                    type="checkbox"
                    checked={Boolean(form.avisoUrgenteAtivo)}
                    onChange={(e) =>
                      setForm({ ...form, avisoUrgenteAtivo: e.target.checked })
                    }
                  />
                  <span>{form.avisoUrgenteAtivo ? "ATIVADO" : "DESATIVADO"}</span>
                </label>
              </div>

              <textarea
                className="emergency-textarea urgent"
                maxLength={160}
                value={form.avisoUrgenteTexto || ""}
                onChange={(e) =>
                  setForm({ ...form, avisoUrgenteTexto: e.target.value })
                }
                placeholder="Ex.: SIMULADO DE ABANDONO ÀS 14H — SIGA AS ORIENTAÇÕES DA BRIGADA."
              />
              <small className="char-counter">
                {(form.avisoUrgenteTexto || "").length}/160
              </small>
            </section>

            {salvo && (
              <div className="admin-save-success">
                ✓ AVISOS SALVOS E ENVIADOS PARA O MURAL
              </div>
            )}

            <button
              type="button"
              onClick={salvarConfig}
              className="save-config-button"
            >
              <Save size={18} /> SALVAR AVISOS E EMERGÊNCIA
            </button>
          </>
        )}

        {aba === "configuracoes" && (
          <section className="painel-card">
            <div className="painel-titulo">
              <Settings size={22} />
              <div>
                <h2>Identidade e unidade</h2>
                <p>Configure os logos e as informações exibidas na Gestão à Vista.</p>
              </div>
            </div>

            {salvo && (
              <div style={{
                padding: 14, marginBottom: 20, borderRadius: 8,
                background: "#e1f7e9", color: "#087c39",
                fontWeight: 900, textAlign: "center",
              }}>
                ✓ CONFIGURAÇÕES SALVAS COM SUCESSO
              </div>
            )}

            <div className="config-logo-grid">
              <div className="config-logo-card">
                <strong>Logo / Brasão da Brigada</strong>
                <div className="config-logo-preview">
                  {form.logoBrigada ? (
                    <img src={form.logoBrigada} alt="Logo da Brigada" />
                  ) : (
                    <ShieldCheck size={60} color="#7c8990" />
                  )}
                </div>
                <input type="file" accept="image/*"
                  onChange={(e) => lerImagemConfig(e.target.files?.[0], "logoBrigada")} />
                {form.logoBrigada && (
                  <button type="button" className="remove-logo"
                    onClick={() => setForm((x) => ({...x, logoBrigada:""}))}>
                    REMOVER LOGO
                  </button>
                )}
              </div>

              <div className="config-logo-card">
                <strong>Logo da Empresa</strong>
                <div className="config-logo-preview">
                  {form.logoEmpresa ? (
                    <img src={form.logoEmpresa} alt="Logo da Empresa" />
                  ) : (
                    <div style={{ color: "#7c8990", fontWeight: 900 }}>LOGO EMPRESA</div>
                  )}
                </div>
                <input type="file" accept="image/*"
                  onChange={(e) => lerImagemConfig(e.target.files?.[0], "logoEmpresa")} />
                {form.logoEmpresa && (
                  <button type="button" className="remove-logo"
                    onClick={() => setForm((x) => ({...x, logoEmpresa:""}))}>
                    REMOVER LOGO
                  </button>
                )}
              </div>
            </div>

            <div className="config-fields-grid">
              <label style={{ fontWeight: 800 }}>
                Nome da Empresa
                <input style={{...campo, marginTop: 7}} value={form.empresaNome}
                  onChange={(e) => setForm({...form, empresaNome:e.target.value})} />
              </label>
              <label style={{ fontWeight: 800 }}>
                Nome da Unidade
                <input style={{...campo, marginTop: 7}} value={form.unidadeNome}
                  onChange={(e) => setForm({...form, unidadeNome:e.target.value})} />
              </label>
              <label style={{ fontWeight: 800 }}>
                Turno
                <input style={{...campo, marginTop: 7}} value={form.turno}
                  onChange={(e) => setForm({...form, turno:e.target.value})} />
              </label>
            </div>

            <div className="geofence-config">
              <div className="geofence-title">
                <MapPin size={22} />
                <div>
                  <h3>Geolocalização da Unidade</h3>
                  <p>
                    Defina o centro da empresa e o raio usado para a presença automática.
                    O painel manual por PIN continuará sempre disponível.
                  </p>
                </div>
              </div>

              <button type="button" className="capture-location" onClick={capturarLocalEmpresa}>
                <Crosshair size={18} /> USAR MINHA LOCALIZAÇÃO ATUAL COMO LOCAL DA EMPRESA
              </button>

              <div className="config-fields-grid geo-fields">
                <label style={{ fontWeight: 800 }}>
                  Latitude
                  <input
                    style={{...campo, marginTop: 7}}
                    value={form.latitude ?? ""}
                    onChange={(e) => setForm({...form, latitude:e.target.value})}
                    placeholder="-23.0000000"
                  />
                </label>

                <label style={{ fontWeight: 800 }}>
                  Longitude
                  <input
                    style={{...campo, marginTop: 7}}
                    value={form.longitude ?? ""}
                    onChange={(e) => setForm({...form, longitude:e.target.value})}
                    placeholder="-46.0000000"
                  />
                </label>

                <label style={{ fontWeight: 800 }}>
                  Raio da unidade (metros)
                  <input
                    style={{...campo, marginTop: 7}}
                    type="number"
                    min="30"
                    max="2000"
                    value={form.raioGeofence ?? 150}
                    onChange={(e) => setForm({...form, raioGeofence:Number(e.target.value)})}
                  />
                </label>

                <label style={{ fontWeight: 800 }}>
                  Precisão máxima aceita (metros)
                  <input
                    style={{...campo, marginTop: 7}}
                    type="number"
                    min="20"
                    max="500"
                    value={form.precisaoMaxima ?? 120}
                    onChange={(e) => setForm({...form, precisaoMaxima:Number(e.target.value)})}
                  />
                </label>
              </div>

              <div className="geofence-rule">
                <b>Regra de segurança:</b> entrada automática ao entrar no raio.
                A saída automática só ocorre para uma presença que tenha sido criada pela
                própria geolocalização. Se o brigadista registrou presença manualmente,
                o GPS não cancela esse registro.
              </div>
            </div>

            <button type="button" onClick={salvarConfig} className="save-config-button">
              SALVAR CONFIGURAÇÕES
            </button>
          </section>
        )}
      </div>
    </div>
  );
}

function App() {
  const [brigadistas, setBrigadistas] =
    useState(carregarDados);
  const [configuracao, setConfiguracao] =
    useState(carregarConfiguracao);
  const [autenticado, setAutenticado] = useState(estaAutenticado);

  useEffect(() => {
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify(brigadistas)
    );
  }, [brigadistas]);

  useEffect(() => {
    localStorage.setItem(CONFIG_KEY, JSON.stringify(configuracao));
  }, [configuracao]);

  const salvarConfiguracao = (novaConfiguracao) => {
    setConfiguracao(novaConfiguracao);
  };

  /*
    IMPORTANTE:
    sincroniza alterações entre duas janelas
    abertas no MESMO computador.
  */
  useEffect(() => {
    const sincronizar = (evento) => {
      if (evento.key === STORAGE_KEY && evento.newValue) {
        try {
          setBrigadistas(JSON.parse(evento.newValue));
        } catch (erro) {
          console.error(erro);
        }
      }

      if (evento.key === CONFIG_KEY && evento.newValue) {
        try {
          setConfiguracao(JSON.parse(evento.newValue));
        } catch (erro) {
          console.error(erro);
        }
      }
    };

    window.addEventListener(
      "storage",
      sincronizar
    );

    return () => {
      window.removeEventListener(
        "storage",
        sincronizar
      );
    };
  }, []);

  const registrarPresenca = (
    id,
    presente,
    origem = "Painel/Terminal"
  ) => {
    setBrigadistas((anteriores) =>
      anteriores.map((pessoa) =>
        pessoa.id === id
          ? {
              ...pessoa,
              presente,
              origemPresenca: origem,
              ultimaAtualizacaoPresenca: new Date().toISOString(),
            }
          : pessoa
      )
    );
  };

  return (
    <Routes>
      <Route
        path="/"
        element={
          <Navigate
            to="/presenca"
            replace
          />
        }
      />

      <Route
        path="/presenca"
        element={
          <Presenca
            brigadistas={brigadistas}
            registrarPresenca={
              registrarPresenca
            }
          />
        }
      />

      <Route
        path="/geolocalizacao"
        element={
          <Geolocalizacao
            brigadistas={brigadistas}
            configuracao={configuracao}
            registrarPresenca={registrarPresenca}
          />
        }
      />

      <Route
        path="/mural"
        element={
          <Mural
            brigadistas={brigadistas}
            configuracao={configuracao}
          />
        }
      />

      <Route
        path="/login"
        element={
          autenticado
            ? <Navigate to="/admin" replace />
            : <Login onLogin={setAutenticado} />
        }
      />

      <Route
        path="/admin"
        element={
          <AdminProtegido autenticado={autenticado}>
            <Admin
              configuracao={configuracao}
              salvarConfiguracao={salvarConfiguracao}
              brigadistas={brigadistas}
              atualizarBrigadistas={setBrigadistas}
              onLogout={() => setAutenticado(false)}
            />
          </AdminProtegido>
        }
      />

      <Route
        path="*"
        element={
          <Navigate
            to="/presenca"
            replace
          />
        }
      />
    </Routes>
  );
}

export default App;