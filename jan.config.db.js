export const dbConfig = {
    server: 'TU_SERVIDOR_SQL',
    authentication: {
        type: 'default',
        options: {
            userName: 'TU_USUARIO',
            password: 'TU_CONTRASEÑA'
        }
    },
    options: {
        encrypt: true,
        database: 'TU_BASE_DE_DATOS',
        trustServerCertificate: true
    }
};