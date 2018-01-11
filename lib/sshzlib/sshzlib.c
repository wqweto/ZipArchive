/*
 * Zlib (RFC1950 / RFC1951) compression for PuTTY.
 * 
 * There will no doubt be criticism of my decision to reimplement
 * Zlib compression from scratch instead of using the existing zlib
 * code. People will cry `reinventing the wheel'; they'll claim
 * that the `fundamental basis of OSS' is code reuse; they'll want
 * to see a really good reason for me having chosen not to use the
 * existing code.
 * 
 * Well, here are my reasons. Firstly, I don't want to link the
 * whole of zlib into the PuTTY binary; PuTTY is justifiably proud
 * of its small size and I think zlib contains a lot of unnecessary
 * baggage for the kind of compression that SSH requires.
 * 
 * Secondly, I also don't like the alternative of using zlib.dll.
 * Another thing PuTTY is justifiably proud of is its ease of
 * installation, and the last thing I want to do is to start
 * mandating DLLs. Not only that, but there are two _kinds_ of
 * zlib.dll kicking around, one with C calling conventions on the
 * exported functions and another with WINAPI conventions, and
 * there would be a significant danger of getting the wrong one.
 * 
 * Thirdly, there seems to be a difference of opinion on the IETF
 * secsh mailing list about the correct way to round off a
 * compressed packet and start the next. In particular, there's
 * some talk of switching to a mechanism zlib isn't currently
 * capable of supporting (see below for an explanation). Given that
 * sort of uncertainty, I thought it might be better to have code
 * that will support even the zlib-incompatible worst case.
 * 
 * Fourthly, it's a _second implementation_. Second implementations
 * are fundamentally a Good Thing in standardisation efforts. The
 * difference of opinion mentioned above has arisen _precisely_
 * because there has been only one zlib implementation and
 * everybody has used it. I don't intend that this should happen
 * again.
 */

#include <stdlib.h>
#include <string.h>
#include <assert.h>

#define ZLIB_STDCALL __stdcall

void ZLIB_STDCALL vbzlib_crc32(struct RelocTable *rtbl, const unsigned char *block, int len, unsigned int *pcrc);
void ZLIB_STDCALL vbzlib_memnonce(unsigned int *block, unsigned int *nonce, int len, int lParam);
void ZLIB_STDCALL vbzlib_memxor(const unsigned char *block, unsigned char *dest, int len, int lParam);
void *ZLIB_STDCALL vbzlib_compress_init(struct RelocTable *rtbl, int wMsg, int wParam, int lParam);
void ZLIB_STDCALL vbzlib_compress_cleanup(void *handle, int wMsg, int wParam, int lParam);
int ZLIB_STDCALL vbzlib_compress_block(void *handle, struct IoBuffers *buf, unsigned int *pcrc, int lParam);
void *ZLIB_STDCALL vbzlib_decompress_init(struct RelocTable *rtbl, int wMsg, int wParam, int lParam);
void ZLIB_STDCALL vbzlib_decompress_cleanup(void *handle, int wMsg, int wParam, int lParam);
int ZLIB_STDCALL vbzlib_decompress_block(void *handle, struct IoBuffers *buf, unsigned int *pcrc, int lParam);
static void ZLIB_STDCALL zlib_literal(struct LZ77Context *ectx, unsigned char c);
static void ZLIB_STDCALL zlib_match(struct LZ77Context *ectx, int distance, int len);

#pragma function(memset)
static void *memset(void *dest, int c, size_t count)
{
    char *bytes = (char *)dest;
    while (count--)
    {
        *bytes++ = (char)c;
    }
    return dest;
}

/*
 * This module also makes a handy zlib decoding tool for when
 * you're picking apart Zip files or PDFs or PNGs. If you compile
 * it with ZLIB_STANDALONE defined, it builds on its own and
 * becomes a command-line utility.
 * 
 * Therefore, here I provide a self-contained implementation of the
 * macros required from the rest of the PuTTY sources.
 */
#define snew(type) ( (type *) rtbl->vbzlib_malloc(sizeof(type)) )
#define snewn(n, type) ( (type *) rtbl->vbzlib_malloc((n) * sizeof(type)) )
#define sresize(x, n, type) ( (type *) rtbl->vbzlib_realloc((x), (n) * sizeof(type)) )
#define sfree(x) ( rtbl->vbzlib_free((x)) )

#ifndef FALSE
#define FALSE 0
#define TRUE (!FALSE)
#endif

struct LZ77Context;
struct LZ77InternalContext;
struct zlib_decompress_ctx;
struct Outbuf;
struct zlib_table;
struct coderecord;

struct RelocTable {
    void *vbzlib_compress_init;
    void *vbzlib_compress_cleanup;
    void *vbzlib_compress_block;
    void *vbzlib_decompress_init;
    void *vbzlib_decompress_cleanup;
    void *vbzlib_decompress_block;
    void *vbzlib_crc32;
    void *vbzlib_memnonce;
    void *vbzlib_memxor;
    void *(ZLIB_STDCALL *vbzlib_malloc)(size_t size);
    void *(ZLIB_STDCALL *vbzlib_realloc)(void *ptr, size_t size);
    void (ZLIB_STDCALL *vbzlib_free)(void *ptr);
    const struct coderecord *lencodes;
    const struct coderecord *distcodes;
    const unsigned char *mirrorbytes;
    const unsigned char *lenlenmap;
    const unsigned int *crc32table;
};

/* ----------------------------------------------------------------------
 * Basic LZ77 code. This bit is designed modularly, so it could be
 * ripped out and used in a different LZ77 compressor. Go to it,
 * and good luck :-)
 */

struct LZ77InternalContext;
struct LZ77Context {
    struct RelocTable *rtbl;
    struct LZ77InternalContext *ictx;
    void *userdata;
};

/*
 * Modifiable parameters.
 */
#define WINSIZE 32768                  /* window size. Must be power of 2! */
#define HASHBITS 16
#define HASHMAX (1<<HASHBITS)          /* one more than max hash value */
#define MAXMATCH 1000                  /* how many matches we track */
#define HASHCHARS 4                    /* how many chars make a hash */

/*
 * This compressor takes a less slapdash approach than the
 * gzip/zlib one. Rather than allowing our hash chains to fall into
 * disuse near the far end, we keep them doubly linked so we can
 * _find_ the far end, and then every time we add a new byte to the
 * window (thus rolling round by one and removing the previous
 * byte), we can carefully remove the hash chain entry.
 */

#define INVALID -1                     /* invalid hash _and_ invalid offset */
struct WindowEntry {
    short next, prev;                  /* array indices within the window */
    int hashval;
};

struct HashEntry {
    short first;                       /* window index of first in chain */
};

struct Match {
    int distance, len;
};

struct LZ77InternalContext {
    struct WindowEntry win[WINSIZE];
    unsigned char data[WINSIZE + 4];
    int winpos;
    struct HashEntry hashtab[HASHMAX];
    unsigned char pending[HASHCHARS];
    int npending;
    int matchdist[MAXMATCH];
};

/*
 * Initialise the private fields of an LZ77Context. It's up to the
 * user to initialise the public fields.
 */
static int ZLIB_STDCALL lz77_init(struct LZ77Context *ctx)
{
    struct RelocTable *rtbl = ctx->rtbl;
    struct LZ77InternalContext *st;
    int i;

    st = snew(struct LZ77InternalContext);
    if (!st)
        return 0;

    ctx->ictx = st;

    for (i = 0; i < WINSIZE; i++)
        st->win[i].next = st->win[i].prev = st->win[i].hashval = INVALID;
    for (i = 0; i < HASHMAX; i++)
        st->hashtab[i].first = INVALID;
    st->winpos = 0;

    st->npending = 0;

    return 1;
}

static void ZLIB_STDCALL lz77_advance(struct LZ77InternalContext *st,
                         unsigned char c, int hash)
{
    int off;

    /*
     * Remove the hash entry at winpos from the tail of its chain,
     * or empty the chain if it's the only thing on the chain.
     */
    if (st->win[st->winpos].prev != INVALID) {
        st->win[st->win[st->winpos].prev].next = INVALID;
    } else if (st->win[st->winpos].hashval != INVALID) {
        st->hashtab[st->win[st->winpos].hashval].first = INVALID;
    }

    /*
     * Create a new entry at winpos and add it to the head of its
     * hash chain.
     */
    st->win[st->winpos].hashval = hash;
    st->win[st->winpos].prev = INVALID;
    off = st->win[st->winpos].next = st->hashtab[hash].first;
    st->hashtab[hash].first = st->winpos;
    if (off != INVALID)
        st->win[off].prev = st->winpos;
    st->data[st->winpos] = c;

    /*
     * Advance the window pointer.
     */
    st->winpos = (st->winpos + 1) & (WINSIZE - 1);
}

#define CHARAT(k) ( (k)<0 ? st->data[(st->winpos+k)&(WINSIZE-1)] : data[k] )
#define DWORDAT(k) *((unsigned int *)&st->data[(st->winpos+k)&(WINSIZE-1)])
// from brotli's Hash frunction
#define kHashMul32 0x1E35A7BD

static __inline int ZLIB_STDCALL lz77_hash3(const unsigned char *data)
{
    return (*((unsigned int *)data) << 8) * kHashMul32 >> (32 - HASHBITS);
}

static __inline int ZLIB_STDCALL lz77_hash(const unsigned char *data)
{
    return *((unsigned int *)data) * kHashMul32 >> (32 - HASHBITS);
}

static void ZLIB_STDCALL lz77_compress_greedy(struct LZ77Context *ctx,
                          const unsigned char *data, int len, int maxmatch, int nicelen)
{
    struct LZ77InternalContext *st = ctx->ictx;
    struct RelocTable *rtbl = ctx->rtbl;
    int i, j, distance, off, nmatch, matchlen, advance, hash;
    int *matches = st->matchdist;

    assert(st->npending <= HASHCHARS);

    /*
     * Add any pending characters from last time to the window. (We
     * might not be able to.)
     *
     * This leaves st->pending empty in the usual case (when len >=
     * HASHCHARS); otherwise it leaves st->pending empty enough that
     * adding all the remaining 'len' characters will not push it past
     * HASHCHARS in size.
     */
    for (i = 0; i < st->npending; i++) {
        unsigned char foo[HASHCHARS];
        if (len + st->npending - i < HASHCHARS) {
            /* Update the pending array. */
            for (j = i; j < st->npending; j++)
                st->pending[j - i] = st->pending[j];
            break;
        }
        for (j = 0; j < HASHCHARS; j++)
            foo[j] = (i + j < st->npending ? st->pending[i + j] :
                      data[i + j - st->npending]);
        lz77_advance(st, foo[0], lz77_hash(foo));
    }
    st->npending -= i;

    while (len > 0) {

        /* Don't even look for a match, if not enough data left. */
        if (len >= HASHCHARS) {
            /*
             * Setup overflow for DWORDAT
             */
            ((unsigned int *)&st->data[WINSIZE])[0] = ((unsigned int *)&st->data[0])[0];

            /*
             * Hash the next few characters.
             */
            hash = lz77_hash(data);

            /*
             * Look the hash up in the corresponding hash chain and see
             * what we can find.
             */
            nmatch = 0;
            for (off = st->hashtab[hash].first;
                 off != INVALID; off = st->win[off].next) {
                /* distance = 1       if off == st->winpos-1 */
                /* distance = WINSIZE if off == st->winpos   */
                distance =
                    WINSIZE - ((off + WINSIZE - st->winpos) & (WINSIZE - 1));
                if (!(((unsigned int *)data)[0] ^ DWORDAT(0 - distance))) {
                    matches[nmatch] = distance;
                    if (++nmatch >= maxmatch)
                        break;
                }
            }
        } else {
            nmatch = 0;
        }

        if (nmatch > 0) {
            /*
             * We've now filled up matches[] with nmatch potential
             * matches. Follow them down to find the longest. (We
             * assume here that it's always worth favouring a
             * longer match over a shorter one.)
             */
            if (nicelen > len)
                nicelen = len;
            
            for (matchlen = HASHCHARS; matchlen < nicelen; matchlen++) {
                unsigned char ch = data[matchlen];
                for (i = j = 0; i < nmatch; i++) {
                    if (ch == CHARAT(matchlen - matches[i])) {
                        matches[j++] = matches[i];
                    }
                }
                if (j <= 1)
                    break;
                nmatch = j;
            }
            while (matchlen < len && data[matchlen] == CHARAT(matchlen - matches[0]))
                matchlen++;

            /*
             * We've now got all the longest matches. We favour the
             * shorter distances, which means we go with matches[0].
             */
            zlib_match(ctx, matches[0], matchlen);
            advance = matchlen;
        } else {
            zlib_literal(ctx, data[0]);
            advance = 1;
        }

        /*
         * Now advance the position by `advance' characters,
         * keeping the window and hash chains consistent.
         */
        while (advance > 0) {
            if (len >= HASHCHARS) {
                lz77_advance(st, *data, lz77_hash(data));
            } else {
                assert(st->npending < HASHCHARS);
                st->pending[st->npending++] = *data;
            }
            data++;
            len--;
            advance--;
        }
    }
}

static void ZLIB_STDCALL lz77_compress_lazy(struct LZ77Context *ctx,
                          const unsigned char *data, int len, int maxmatch, int nicelen)
{
    struct LZ77InternalContext *st = ctx->ictx;
    struct RelocTable *rtbl = ctx->rtbl;
    int i, j, distance, off, nmatch, matchlen, advance, hash;
    struct Match defermatch;
    int *matches = st->matchdist;
    int deferchr;

    assert(st->npending <= HASHCHARS);

    /*
     * Add any pending characters from last time to the window. (We
     * might not be able to.)
     *
     * This leaves st->pending empty in the usual case (when len >=
     * HASHCHARS); otherwise it leaves st->pending empty enough that
     * adding all the remaining 'len' characters will not push it past
     * HASHCHARS in size.
     */
    for (i = 0; i < st->npending; i++) {
        unsigned char foo[HASHCHARS];
        if (len + st->npending - i < HASHCHARS) {
            /* Update the pending array. */
            for (j = i; j < st->npending; j++)
                st->pending[j - i] = st->pending[j];
            break;
        }
        for (j = 0; j < HASHCHARS; j++)
            foo[j] = (i + j < st->npending ? st->pending[i + j] :
                      data[i + j - st->npending]);
        lz77_advance(st, foo[0], lz77_hash(foo));
    }
    st->npending -= i;

    defermatch.distance = 0; /* appease compiler */
    defermatch.len = 0;
    deferchr = '\0';
    while (len > 0) {

        /* Don't even look for a match, if not enough data left. */
        if (len >= HASHCHARS) {
            /*
             * Setup overflow for DWORDAT
             */
            ((unsigned int *)&st->data[WINSIZE])[0] = ((unsigned int *)&st->data[0])[0];

            /*
             * Hash the next few characters.
             */
            hash = lz77_hash(data);

            /*
             * Look the hash up in the corresponding hash chain and see
             * what we can find.
             */
            nmatch = 0;
            for (off = st->hashtab[hash].first;
                 off != INVALID; off = st->win[off].next) {
                /* distance = 1       if off == st->winpos-1 */
                /* distance = WINSIZE if off == st->winpos   */
                distance =
                    WINSIZE - ((off + WINSIZE - st->winpos) & (WINSIZE - 1));
                if (!(((unsigned int *)data)[0] ^ DWORDAT(0 - distance))) {
                    matches[nmatch] = distance;
                    if (++nmatch >= maxmatch)
                        break;
                }
            }
        } else {
            nmatch = 0;
        }

        if (nmatch > 0) {
            /*
             * We've now filled up matches[] with nmatch potential
             * matches. Follow them down to find the longest. (We
             * assume here that it's always worth favouring a
             * longer match over a shorter one.)
             */
            if (nicelen > len)
                nicelen = len;
            
            for (matchlen = HASHCHARS; matchlen < nicelen; matchlen++) {
                unsigned char ch = data[matchlen];
                for (i = j = 0; i < nmatch; i++) {
                    if (ch == CHARAT(matchlen - matches[i])) {
                        matches[j++] = matches[i];
                    }
                }
                if (j <= 1)
                    break;
                nmatch = j;
            }
            while (matchlen < len && data[matchlen] == CHARAT(matchlen - matches[0]))
                matchlen++;

            /*
             * We've now got all the longest matches. We favour the
             * shorter distances, which means we go with matches[0].
             * So see if we want to defer it or throw it away.
             */
            if (defermatch.len > 0) {
                if (matchlen > defermatch.len + 1) {
                    /* We have a better match. Emit the deferred char,
                     * and defer this match. */
                    zlib_literal(ctx, (unsigned char) deferchr);
                    defermatch.distance = matches[0];
                    defermatch.len = matchlen;
                    deferchr = data[0];
                    advance = 1;
                } else {
                    /* We don't have a better match. Do the deferred one. */
                    zlib_match(ctx, defermatch.distance, defermatch.len);
                    advance = defermatch.len - 1;
                    defermatch.len = 0;
                }
            } else {
                /* There was no deferred match. Defer this one. */
                defermatch.distance = matches[0];
                defermatch.len = matchlen;
                deferchr = data[0];
                advance = 1;
            }
        } else {
            /*
             * We found no matches. Emit the deferred match, if
             * any; otherwise emit a literal.
             */
            if (defermatch.len > 0) {
                zlib_match(ctx, defermatch.distance, defermatch.len);
                advance = defermatch.len - 1;
                defermatch.len = 0;
            } else {
                zlib_literal(ctx, data[0]);
                advance = 1;
            }
        }

        /*
         * Now advance the position by `advance' characters,
         * keeping the window and hash chains consistent.
         */
        while (advance > 0) {
            if (len >= HASHCHARS) {
                lz77_advance(st, *data, lz77_hash(data));
            } else {
                assert(st->npending < HASHCHARS);
                st->pending[st->npending++] = *data;
            }
            data++;
            len--;
            advance--;
        }
    }
}

#undef HASHCHARS
#define HASHCHARS 3
#define lz77_hash lz77_hash3

static void ZLIB_STDCALL lz77_compress_lazy_hash3(struct LZ77Context *ctx,
                          const unsigned char *data, int len)
{
    struct LZ77InternalContext *st = ctx->ictx;
    struct RelocTable *rtbl = ctx->rtbl;
    int i, j, distance, off, nmatch, matchlen, advance, hash;
    struct Match defermatch;
    int *matches = st->matchdist;
    int deferchr;

    assert(st->npending <= HASHCHARS);

    /*
     * Add any pending characters from last time to the window. (We
     * might not be able to.)
     *
     * This leaves st->pending empty in the usual case (when len >=
     * HASHCHARS); otherwise it leaves st->pending empty enough that
     * adding all the remaining 'len' characters will not push it past
     * HASHCHARS in size.
     */
    for (i = 0; i < st->npending; i++) {
        unsigned char foo[HASHCHARS];
        if (len + st->npending - i < HASHCHARS) {
            /* Update the pending array. */
            for (j = i; j < st->npending; j++)
                st->pending[j - i] = st->pending[j];
            break;
        }
        for (j = 0; j < HASHCHARS; j++)
            foo[j] = (i + j < st->npending ? st->pending[i + j] :
                      data[i + j - st->npending]);
        lz77_advance(st, foo[0], lz77_hash(foo));
    }
    st->npending -= i;

    defermatch.distance = 0; /* appease compiler */
    defermatch.len = 0;
    deferchr = '\0';
    while (len > 0) {

        /* Don't even look for a match, if not enough data left. */
        if (len >= HASHCHARS) {
            /*
             * Hash the next few characters.
             */
            hash = lz77_hash(data);

            /*
             * Look the hash up in the corresponding hash chain and see
             * what we can find.
             */
            nmatch = 0;
            for (off = st->hashtab[hash].first;
                 off != INVALID; off = st->win[off].next) {
                /* distance = 1       if off == st->winpos-1 */
                /* distance = WINSIZE if off == st->winpos   */
                distance =
                    WINSIZE - ((off + WINSIZE - st->winpos) & (WINSIZE - 1));
                if (data[0] == CHARAT(0 - distance)
                        && data[1] == CHARAT(1 - distance)
                        && data[2] == CHARAT(2 - distance)) {
                    matches[nmatch] = distance;
                    if (++nmatch >= MAXMATCH)
                        break;
                }
            }
        } else {
            nmatch = 0;
        }

        if (nmatch > 0) {
            /*
             * We've now filled up matches[] with nmatch potential
             * matches. Follow them down to find the longest. (We
             * assume here that it's always worth favouring a
             * longer match over a shorter one.)
             */
            for (matchlen = HASHCHARS; matchlen < len; matchlen++) {
                unsigned char ch = data[matchlen];
                for (i = j = 0; i < nmatch; i++) {
                    if (ch == CHARAT(matchlen - matches[i])) {
                        matches[j++] = matches[i];
                    }
                }
                if (j <= 1)
                    break;
                nmatch = j;
            }
            while (matchlen < len && data[matchlen] == CHARAT(matchlen - matches[0]))
                matchlen++;

            /*
             * We've now got all the longest matches. We favour the
             * shorter distances, which means we go with matches[0].
             * So see if we want to defer it or throw it away.
             */
            if (defermatch.len > 0) {
                if (matchlen > defermatch.len + 1) {
                    /* We have a better match. Emit the deferred char,
                     * and defer this match. */
                    zlib_literal(ctx, (unsigned char) deferchr);
                    defermatch.distance = matches[0];
                    defermatch.len = matchlen;
                    deferchr = data[0];
                    advance = 1;
                } else {
                    /* We don't have a better match. Do the deferred one. */
                    zlib_match(ctx, defermatch.distance, defermatch.len);
                    advance = defermatch.len - 1;
                    defermatch.len = 0;
                }
            } else {
                /* There was no deferred match. Defer this one. */
                defermatch.distance = matches[0];
                defermatch.len = matchlen;
                deferchr = data[0];
                advance = 1;
            }
        } else {
            /*
             * We found no matches. Emit the deferred match, if
             * any; otherwise emit a literal.
             */
            if (defermatch.len > 0) {
                zlib_match(ctx, defermatch.distance, defermatch.len);
                advance = defermatch.len - 1;
                defermatch.len = 0;
            } else {
                zlib_literal(ctx, data[0]);
                advance = 1;
            }
        }

        /*
         * Now advance the position by `advance' characters,
         * keeping the window and hash chains consistent.
         */
        while (advance > 0) {
            if (len >= HASHCHARS) {
                lz77_advance(st, *data, lz77_hash(data));
            } else {
                assert(st->npending < HASHCHARS);
                st->pending[st->npending++] = *data;
            }
            data++;
            len--;
            advance--;
        }
    }
}

/* ----------------------------------------------------------------------
 * Zlib compression. We always use the static Huffman tree option.
 * Mostly this is because it's hard to scan a block in advance to
 * work out better trees; dynamic trees are great when you're
 * compressing a large file under no significant time constraint,
 * but when you're compressing little bits in real time, things get
 * hairier.
 * 
 * I suppose it's possible that I could compute Huffman trees based
 * on the frequencies in the _previous_ block, as a sort of
 * heuristic, but I'm not confident that the gain would balance out
 * having to transmit the trees.
 */

struct Outbuf {
    unsigned char *outbuf;
    int outlen, outsize;
    unsigned long outbits;
    int noutbits;
};

static void ZLIB_STDCALL zlib_outbits(struct RelocTable *rtbl, struct Outbuf *out, unsigned long bits, int nbits)
{
    assert(out->noutbits + nbits <= 32);
    out->outbits |= bits << out->noutbits;
    out->noutbits += nbits;
    while (out->noutbits >= 8) {
        if (out->outlen >= out->outsize) {
            out->outsize = out->outsize << 1;
            out->outbuf = sresize(out->outbuf, out->outsize, unsigned char);
        }
        out->outbuf[out->outlen++] = (unsigned char) (out->outbits & 0xFF);
        out->outbits >>= 8;
        out->noutbits -= 8;
    }
}

static const unsigned char zdat_mirrorbytes[256] = {
    0x00, 0x80, 0x40, 0xc0, 0x20, 0xa0, 0x60, 0xe0,
    0x10, 0x90, 0x50, 0xd0, 0x30, 0xb0, 0x70, 0xf0,
    0x08, 0x88, 0x48, 0xc8, 0x28, 0xa8, 0x68, 0xe8,
    0x18, 0x98, 0x58, 0xd8, 0x38, 0xb8, 0x78, 0xf8,
    0x04, 0x84, 0x44, 0xc4, 0x24, 0xa4, 0x64, 0xe4,
    0x14, 0x94, 0x54, 0xd4, 0x34, 0xb4, 0x74, 0xf4,
    0x0c, 0x8c, 0x4c, 0xcc, 0x2c, 0xac, 0x6c, 0xec,
    0x1c, 0x9c, 0x5c, 0xdc, 0x3c, 0xbc, 0x7c, 0xfc,
    0x02, 0x82, 0x42, 0xc2, 0x22, 0xa2, 0x62, 0xe2,
    0x12, 0x92, 0x52, 0xd2, 0x32, 0xb2, 0x72, 0xf2,
    0x0a, 0x8a, 0x4a, 0xca, 0x2a, 0xaa, 0x6a, 0xea,
    0x1a, 0x9a, 0x5a, 0xda, 0x3a, 0xba, 0x7a, 0xfa,
    0x06, 0x86, 0x46, 0xc6, 0x26, 0xa6, 0x66, 0xe6,
    0x16, 0x96, 0x56, 0xd6, 0x36, 0xb6, 0x76, 0xf6,
    0x0e, 0x8e, 0x4e, 0xce, 0x2e, 0xae, 0x6e, 0xee,
    0x1e, 0x9e, 0x5e, 0xde, 0x3e, 0xbe, 0x7e, 0xfe,
    0x01, 0x81, 0x41, 0xc1, 0x21, 0xa1, 0x61, 0xe1,
    0x11, 0x91, 0x51, 0xd1, 0x31, 0xb1, 0x71, 0xf1,
    0x09, 0x89, 0x49, 0xc9, 0x29, 0xa9, 0x69, 0xe9,
    0x19, 0x99, 0x59, 0xd9, 0x39, 0xb9, 0x79, 0xf9,
    0x05, 0x85, 0x45, 0xc5, 0x25, 0xa5, 0x65, 0xe5,
    0x15, 0x95, 0x55, 0xd5, 0x35, 0xb5, 0x75, 0xf5,
    0x0d, 0x8d, 0x4d, 0xcd, 0x2d, 0xad, 0x6d, 0xed,
    0x1d, 0x9d, 0x5d, 0xdd, 0x3d, 0xbd, 0x7d, 0xfd,
    0x03, 0x83, 0x43, 0xc3, 0x23, 0xa3, 0x63, 0xe3,
    0x13, 0x93, 0x53, 0xd3, 0x33, 0xb3, 0x73, 0xf3,
    0x0b, 0x8b, 0x4b, 0xcb, 0x2b, 0xab, 0x6b, 0xeb,
    0x1b, 0x9b, 0x5b, 0xdb, 0x3b, 0xbb, 0x7b, 0xfb,
    0x07, 0x87, 0x47, 0xc7, 0x27, 0xa7, 0x67, 0xe7,
    0x17, 0x97, 0x57, 0xd7, 0x37, 0xb7, 0x77, 0xf7,
    0x0f, 0x8f, 0x4f, 0xcf, 0x2f, 0xaf, 0x6f, 0xef,
    0x1f, 0x9f, 0x5f, 0xdf, 0x3f, 0xbf, 0x7f, 0xff,
};

struct coderecord {
    short code, extrabits;
    int min, max;
};

static const struct coderecord zdat_lencodes[] = {
    {257, 0, 3, 3},
    {258, 0, 4, 4},
    {259, 0, 5, 5},
    {260, 0, 6, 6},
    {261, 0, 7, 7},
    {262, 0, 8, 8},
    {263, 0, 9, 9},
    {264, 0, 10, 10},
    {265, 1, 11, 12},
    {266, 1, 13, 14},
    {267, 1, 15, 16},
    {268, 1, 17, 18},
    {269, 2, 19, 22},
    {270, 2, 23, 26},
    {271, 2, 27, 30},
    {272, 2, 31, 34},
    {273, 3, 35, 42},
    {274, 3, 43, 50},
    {275, 3, 51, 58},
    {276, 3, 59, 66},
    {277, 4, 67, 82},
    {278, 4, 83, 98},
    {279, 4, 99, 114},
    {280, 4, 115, 130},
    {281, 5, 131, 162},
    {282, 5, 163, 194},
    {283, 5, 195, 226},
    {284, 5, 227, 257},
    {285, 0, 258, 258},
};

static const struct coderecord zdat_distcodes[] = {
    {0, 0, 1, 1},
    {1, 0, 2, 2},
    {2, 0, 3, 3},
    {3, 0, 4, 4},
    {4, 1, 5, 6},
    {5, 1, 7, 8},
    {6, 2, 9, 12},
    {7, 2, 13, 16},
    {8, 3, 17, 24},
    {9, 3, 25, 32},
    {10, 4, 33, 48},
    {11, 4, 49, 64},
    {12, 5, 65, 96},
    {13, 5, 97, 128},
    {14, 6, 129, 192},
    {15, 6, 193, 256},
    {16, 7, 257, 384},
    {17, 7, 385, 512},
    {18, 8, 513, 768},
    {19, 8, 769, 1024},
    {20, 9, 1025, 1536},
    {21, 9, 1537, 2048},
    {22, 10, 2049, 3072},
    {23, 10, 3073, 4096},
    {24, 11, 4097, 6144},
    {25, 11, 6145, 8192},
    {26, 12, 8193, 12288},
    {27, 12, 12289, 16384},
    {28, 13, 16385, 24576},
    {29, 13, 24577, 32768},
};

static const unsigned char zdat_lenlenmap[] = {
    16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15
};

static void ZLIB_STDCALL zlib_literal(struct LZ77Context *ectx, unsigned char c)
{
    struct RelocTable *rtbl = ectx->rtbl;
    struct Outbuf *out = (struct Outbuf *) ectx->userdata;

    if (c <= 143) {
        /* 0 through 143 are 8 bits long starting at 00110000. */
        zlib_outbits(rtbl, out, rtbl->mirrorbytes[0x30 + c], 8);
    } else {
        /* 144 through 255 are 9 bits long starting at 110010000. */
        zlib_outbits(rtbl, out, 1 + 2 * rtbl->mirrorbytes[0x90 - 144 + c], 9);
    }
}

static void ZLIB_STDCALL zlib_match(struct LZ77Context *ectx, int distance, int len)
{
    struct RelocTable *rtbl = ectx->rtbl;
    const struct coderecord *d, *l;
    int i, j, k;
    struct Outbuf *out = (struct Outbuf *) ectx->userdata;
    
    while (len > 0) {
        int thislen;

        /*
         * We can transmit matches of lengths 3 through 258
         * inclusive. So if len exceeds 258, we must transmit in
         * several steps, with 258 or less in each step.
         * 
         * Specifically: if len >= 261, we can transmit 258 and be
         * sure of having at least 3 left for the next step. And if
         * len <= 258, we can just transmit len. But if len == 259
         * or 260, we must transmit len-3.
         */
        thislen = (len > 260 ? 258 : len <= 258 ? len : len - 3);
        len -= thislen;

        /*
         * Binary-search to find which length code we're
         * transmitting.
         */
        i = -1;
        j = sizeof(zdat_lencodes) / sizeof(*zdat_lencodes);
        while (1) {
            assert(j - i >= 2);
            k = (j + i) / 2;
            if (thislen < rtbl->lencodes[k].min)
                j = k;
            else if (thislen > rtbl->lencodes[k].max)
                i = k;
            else {
                l = &rtbl->lencodes[k];
                break;                 /* found it! */
            }
        }

        /*
         * Transmit the length code. 256-279 are seven bits
         * starting at 0000000; 280-287 are eight bits starting at
         * 11000000.
         */
        if (l->code <= 279) {
            zlib_outbits(rtbl, out, rtbl->mirrorbytes[(l->code - 256) * 2], 7);
        } else {
            zlib_outbits(rtbl, out, rtbl->mirrorbytes[0xc0 - 280 + l->code], 8);
        }

        /*
         * Transmit the extra bits.
         */
        if (l->extrabits)
            zlib_outbits(rtbl, out, thislen - l->min, l->extrabits);

        /*
         * Binary-search to find which distance code we're
         * transmitting.
         */
        i = -1;
        j = sizeof(zdat_distcodes) / sizeof(*zdat_distcodes);
        while (1) {
            assert(j - i >= 2);
            k = (j + i) / 2;
            if (distance < rtbl->distcodes[k].min)
                j = k;
            else if (distance > rtbl->distcodes[k].max)
                i = k;
            else {
                d = &rtbl->distcodes[k];
                break;                 /* found it! */
            }
        }

        /*
         * Transmit the distance code. Five bits starting at 00000.
         */
        zlib_outbits(rtbl, out, rtbl->mirrorbytes[d->code * 8], 5);

        /*
         * Transmit the extra bits.
         */
        if (d->extrabits)
            zlib_outbits(rtbl, out, distance - d->min, d->extrabits);
    }
}

void *ZLIB_STDCALL zlib_compress_init(struct RelocTable *rtbl)
{
    struct Outbuf *out;
    struct LZ77Context *ectx = snew(struct LZ77Context);

    ectx->rtbl = rtbl;
    lz77_init(ectx);

    out = snew(struct Outbuf);
    out->outbits = out->noutbits = 0;
    ectx->userdata = out;

    return ectx;
}

void ZLIB_STDCALL zlib_compress_cleanup(void *handle)
{
    struct LZ77Context *ectx = (struct LZ77Context *)handle;
    struct RelocTable *rtbl = ectx->rtbl;
    sfree(ectx->userdata);
    sfree(ectx->ictx);
    sfree(ectx);
}

int ZLIB_STDCALL zlib_compress_block(void *handle, const unsigned char *block, int len,
                        unsigned char **outblock, int *outlen, int final, int greedy, int maxmatch, int nicelen)
{
    struct LZ77Context *ectx = (struct LZ77Context *)handle;
    struct RelocTable *rtbl = ectx->rtbl;
    struct Outbuf *out = (struct Outbuf *) ectx->userdata;

    out->outbuf = snewn(32768, unsigned char);
    out->outsize = 32768;
    out->outlen = 0;

    /*
     * Start a Deflate (RFC1951) fixed-trees block. We
     * transmit a bit (BFINAL=final), followed by a zero
     * bit and a one bit (BTYPE=01). Of course these are in
     * the wrong order (01 final).
     */
    zlib_outbits(rtbl, out, 2 + (final != 0), 3);

    /*
     * Do the compression.
     */
    if (greedy)
        lz77_compress_greedy(ectx, block, len, maxmatch, nicelen);
    else {
#ifdef ZLIB_COMPRESS_LAZY_HASH3
        if (maxmatch < 200)
            lz77_compress_lazy(ectx, block, len, maxmatch, nicelen);
        else
            lz77_compress_lazy_hash3(ectx, block, len);
#else
        lz77_compress_lazy(ectx, block, len, maxmatch, nicelen);
#endif
    }
    zlib_outbits(rtbl, out, 0, 7);            /* close block */
    if (final && out->noutbits)               /* align to a byte boundary */
        zlib_outbits(rtbl, out, 0, 8 - out->noutbits); 

    *outblock = out->outbuf;
    *outlen = out->outlen;

    return 1;
}

/* ----------------------------------------------------------------------
 * Zlib decompression. Of course, even though our compressor always
 * uses static trees, our _decompressor_ has to be capable of
 * handling dynamic trees if it sees them.
 */

/*
 * The way we work the Huffman decode is to have a table lookup on
 * the first N bits of the input stream (in the order they arrive,
 * of course, i.e. the first bit of the Huffman code is in bit 0).
 * Each table entry lists the number of bits to consume, plus
 * either an output code or a pointer to a secondary table.
 */
struct zlib_table;
struct zlib_tableentry;

struct zlib_tableentry {
    unsigned char nbits;
    short code;
    struct zlib_table *nexttable;
};

struct zlib_table {
    int mask;                          /* mask applied to input bit stream */
    struct zlib_tableentry *table;
};

#define MAXCODELEN 16
#define MAXSYMS 288

/*
 * Build a single-level decode table for elements
 * [minlength,maxlength) of the provided code/length tables, and
 * recurse to build subtables.
 */
static struct zlib_table *ZLIB_STDCALL zlib_mkonetab(struct RelocTable *rtbl, int *codes, unsigned char *lengths,
                                        int nsyms,
                                        int pfx, int pfxbits, int bits)
{
    struct zlib_table *tab = snew(struct zlib_table);
    int pfxmask = (1 << pfxbits) - 1;
    int nbits, i, j, code;

    tab->table = snewn(1 << bits, struct zlib_tableentry);
    tab->mask = (1 << bits) - 1;

    for (code = 0; code <= tab->mask; code++) {
        tab->table[code].code = -1;
        tab->table[code].nbits = 0;
        tab->table[code].nexttable = NULL;
    }

    for (i = 0; i < nsyms; i++) {
        if (lengths[i] <= pfxbits || (codes[i] & pfxmask) != pfx)
            continue;
        code = (codes[i] >> pfxbits) & tab->mask;
        for (j = code; j <= tab->mask; j += 1 << (lengths[i] - pfxbits)) {
            tab->table[j].code = i;
            nbits = lengths[i] - pfxbits;
            if (tab->table[j].nbits < nbits)
                tab->table[j].nbits = nbits;
        }
    }
    for (code = 0; code <= tab->mask; code++) {
        if (tab->table[code].nbits <= bits)
            continue;
        /* Generate a subtable. */
        tab->table[code].code = -1;
        nbits = tab->table[code].nbits - bits;
        if (nbits > 7)
            nbits = 7;
        tab->table[code].nbits = bits;
        tab->table[code].nexttable = zlib_mkonetab(rtbl, codes, lengths, nsyms,
                                                   pfx | (code << pfxbits),
                                                   pfxbits + bits, nbits);
    }

    return tab;
}

struct zlib_decompress_ctx {
    struct RelocTable *rtbl;
    struct zlib_table *staticlentable, *staticdisttable;
    struct zlib_table *currlentable, *currdisttable, *lenlentable;
    enum {
        START, OUTSIDEBLK,
        TREES_HDR, TREES_LENLEN, TREES_LEN, TREES_LENREP,
        INBLK, GOTLENSYM, GOTLEN, GOTDISTSYM,
        UNCOMP_LEN, UNCOMP_NLEN, UNCOMP_DATA
    } state;
    int sym, hlit, hdist, hclen, lenptr, lenextrabits, lenaddon, len,
        lenrep;
    int uncomplen;
    unsigned char lenlen[19];
    unsigned char lengths[286 + 32];
    unsigned long bits;
    int nbits;
    unsigned char window[WINSIZE];
    int winpos;
    unsigned char *outblk;
    int outlen, outsize;
};

/*
 * Build a decode table, given a set of Huffman tree lengths.
 */
static struct zlib_table *ZLIB_STDCALL zlib_mktable(struct zlib_decompress_ctx *dctx, unsigned char *lengths,
                                       int nlengths)
{
    int count[MAXCODELEN], startcode[MAXCODELEN], codes[MAXSYMS];
    int code, maxlen;
    int i, j;

    /* Count the codes of each length. */
    maxlen = 0;
    for (i = 1; i < MAXCODELEN; i++)
        count[i] = 0;
    for (i = 0; i < nlengths; i++) {
        count[lengths[i]]++;
        if (maxlen < lengths[i])
            maxlen = lengths[i];
    }
    /* Determine the starting code for each length block. */
    code = 0;
    for (i = 1; i < MAXCODELEN; i++) {
        startcode[i] = code;
        code += count[i];
        code <<= 1;
    }
    /* Determine the code for each symbol. Mirrored, of course. */
    for (i = 0; i < nlengths; i++) {
        code = startcode[lengths[i]]++;
        codes[i] = 0;
        for (j = 0; j < lengths[i]; j++) {
            codes[i] = (codes[i] << 1) | (code & 1);
            code >>= 1;
        }
    }

    /*
     * Now we have the complete list of Huffman codes. Build a
     * table.
     */
    return zlib_mkonetab(dctx->rtbl, codes, lengths, nlengths, 0, 0,
                         maxlen < 9 ? maxlen : 9);
}

static int ZLIB_STDCALL zlib_freetable(struct zlib_decompress_ctx *dctx, struct zlib_table **ztab)
{
    struct RelocTable *rtbl = dctx->rtbl;
    struct zlib_table *tab;
    int code;

    if (ztab == NULL)
        return -1;

    if (*ztab == NULL)
        return 0;

    tab = *ztab;

    for (code = 0; code <= tab->mask; code++)
        if (tab->table[code].nexttable != NULL)
            zlib_freetable(dctx, &tab->table[code].nexttable);

    sfree(tab->table);
    tab->table = NULL;

    sfree(tab);
    *ztab = NULL;

    return (0);
}

void *ZLIB_STDCALL zlib_decompress_init(struct RelocTable *rtbl)
{
    struct zlib_decompress_ctx *dctx = snew(struct zlib_decompress_ctx);
    unsigned char lengths[288];

    dctx->rtbl = rtbl;
    memset(lengths, 8, 144);
    memset(lengths + 144, 9, 256 - 144);
    memset(lengths + 256, 7, 280 - 256);
    memset(lengths + 280, 8, 288 - 280);
    dctx->staticlentable = zlib_mktable(dctx, lengths, 288);
    memset(lengths, 5, 32);
    dctx->staticdisttable = zlib_mktable(dctx, lengths, 32);
    dctx->state = START;                       /* even before header */
    dctx->currlentable = dctx->currdisttable = dctx->lenlentable = NULL;
    dctx->bits = 0;
    dctx->nbits = 0;
    dctx->winpos = 0;

    return dctx;
}

void ZLIB_STDCALL zlib_decompress_cleanup(void *handle)
{
    struct zlib_decompress_ctx *dctx = (struct zlib_decompress_ctx *)handle;
    struct RelocTable *rtbl = dctx->rtbl;

    if (dctx->currlentable && dctx->currlentable != dctx->staticlentable)
        zlib_freetable(dctx, &dctx->currlentable);
    if (dctx->currdisttable && dctx->currdisttable != dctx->staticdisttable)
        zlib_freetable(dctx, &dctx->currdisttable);
    if (dctx->lenlentable)
        zlib_freetable(dctx, &dctx->lenlentable);
    zlib_freetable(dctx, &dctx->staticlentable);
    zlib_freetable(dctx, &dctx->staticdisttable);
    sfree(dctx);
}

static int ZLIB_STDCALL zlib_huflookup(unsigned long *bitsp, int *nbitsp,
                   struct zlib_table *tab)
{
    unsigned long bits = *bitsp;
    int nbits = *nbitsp;
    while (1) {
        struct zlib_tableentry *ent;
        ent = &tab->table[bits & tab->mask];
        if (ent->nbits > nbits)
            return -1;                 /* not enough data */
        bits >>= ent->nbits;
        nbits -= ent->nbits;
        if (ent->code == -1)
            tab = ent->nexttable;
        else {
            *bitsp = bits;
            *nbitsp = nbits;
            return ent->code;
        }

        if (!tab) {
            /*
             * There was a missing entry in the table, presumably
             * due to an invalid Huffman table description, and the
             * subsequent data has attempted to use the missing
             * entry. Return a decoding failure.
             */
            return -2;
        }
    }
}

static void ZLIB_STDCALL zlib_emit_char(struct zlib_decompress_ctx *dctx, int c)
{
    struct RelocTable *rtbl = dctx->rtbl;
    
    dctx->window[dctx->winpos] = c;
    dctx->winpos = (dctx->winpos + 1) & (WINSIZE - 1);
    if (dctx->outlen >= dctx->outsize) {
        dctx->outsize = dctx->outlen << 1;
        dctx->outblk = sresize(dctx->outblk, dctx->outsize, unsigned char);
    }
    dctx->outblk[dctx->outlen++] = c;
}

#define EATBITS(n) ( dctx->nbits -= (n), dctx->bits >>= (n) )

int ZLIB_STDCALL zlib_decompress_block(void *handle, const unsigned char *block, int len,
                          unsigned char **outblock, int *outlen)
{
    struct zlib_decompress_ctx *dctx = (struct zlib_decompress_ctx *)handle;
    struct RelocTable *rtbl = dctx->rtbl;
    const struct coderecord *rec;
    int code, blktype, rep, dist, nlen, header;

    dctx->outblk = snewn(32768, unsigned char);
    dctx->outsize = 32768;
    dctx->outlen = 0;

    while (len > 0 || dctx->nbits > 0) {
        while (dctx->nbits < 24 && len > 0) {
            dctx->bits |= (*block++) << dctx->nbits;
            dctx->nbits += 8;
            len--;
        }
        if (dctx->state == START) {
            dctx->state = OUTSIDEBLK;
        }
        else if (dctx->state == OUTSIDEBLK) {
            /* Expect 3-bit block header. */
            if (dctx->nbits < 3)
                goto finished;         /* done all we can */
            EATBITS(1);
            blktype = dctx->bits & 3;
            EATBITS(2);
            if (blktype == 0) {
                int to_eat = dctx->nbits & 7;
                dctx->state = UNCOMP_LEN;
                EATBITS(to_eat);       /* align to byte boundary */
            } else if (blktype == 1) {
                dctx->currlentable = dctx->staticlentable;
                dctx->currdisttable = dctx->staticdisttable;
                dctx->state = INBLK;
            } else if (blktype == 2) {
                dctx->state = TREES_HDR;
            }
        }
        else if (dctx->state == TREES_HDR) {
            /*
             * Dynamic block header. Five bits of HLIT, five of
             * HDIST, four of HCLEN.
             */
            if (dctx->nbits < 5 + 5 + 4)
                goto finished;         /* done all we can */
            dctx->hlit = 257 + (dctx->bits & 31);
            EATBITS(5);
            dctx->hdist = 1 + (dctx->bits & 31);
            EATBITS(5);
            dctx->hclen = 4 + (dctx->bits & 15);
            EATBITS(4);
            dctx->lenptr = 0;
            dctx->state = TREES_LENLEN;
            memset(dctx->lenlen, 0, sizeof(dctx->lenlen));
        }
        else if (dctx->state == TREES_LENLEN) {
            if (dctx->nbits < 3)
                goto finished;
            while (dctx->lenptr < dctx->hclen && dctx->nbits >= 3) {
                dctx->lenlen[rtbl->lenlenmap[dctx->lenptr++]] =
                    (unsigned char) (dctx->bits & 7);
                EATBITS(3);
            }
            if (dctx->lenptr == dctx->hclen) {
                dctx->lenlentable = zlib_mktable(dctx, dctx->lenlen, 19);
                dctx->state = TREES_LEN;
                dctx->lenptr = 0;
            }
        }
        else if (dctx->state == TREES_LEN) {
            if (dctx->lenptr >= dctx->hlit + dctx->hdist) {
                dctx->currlentable = zlib_mktable(dctx, dctx->lengths, dctx->hlit);
                dctx->currdisttable = zlib_mktable(dctx, dctx->lengths + dctx->hlit,
                                                  dctx->hdist);
                zlib_freetable(dctx, &dctx->lenlentable);
                dctx->lenlentable = NULL;
                dctx->state = INBLK;
            }
            else {
                code =
                    zlib_huflookup(&dctx->bits, &dctx->nbits, dctx->lenlentable);
                if (code == -1)
                    goto finished;
                if (code == -2)
                    goto decode_error;
                if (code < 16)
                    dctx->lengths[dctx->lenptr++] = code;
                else {
                    dctx->lenextrabits = (code == 16 ? 2 : code == 17 ? 3 : 7);
                    dctx->lenaddon = (code == 18 ? 11 : 3);
                    dctx->lenrep = (code == 16 && dctx->lenptr > 0 ?
                                   dctx->lengths[dctx->lenptr - 1] : 0);
                    dctx->state = TREES_LENREP;
                }
            }
        }
        else if (dctx->state == TREES_LENREP) {
            if (dctx->nbits < dctx->lenextrabits)
                goto finished;
            rep =
                dctx->lenaddon +
                (dctx->bits & ((1 << dctx->lenextrabits) - 1));
            EATBITS(dctx->lenextrabits);
            while (rep > 0 && dctx->lenptr < dctx->hlit + dctx->hdist) {
                dctx->lengths[dctx->lenptr] = dctx->lenrep;
                dctx->lenptr++;
                rep--;
            }
            dctx->state = TREES_LEN;
        }
        else if (dctx->state == INBLK) {
            code =
                zlib_huflookup(&dctx->bits, &dctx->nbits, dctx->currlentable);
            if (code == -1)
                goto finished;
            if (code == -2)
                goto decode_error;
            if (code < 256)
                zlib_emit_char(dctx, code);
            else if (code == 256) {
                dctx->state = OUTSIDEBLK;
                if (dctx->currlentable != dctx->staticlentable) {
                    zlib_freetable(dctx, &dctx->currlentable);
                    dctx->currlentable = NULL;
                }
                if (dctx->currdisttable != dctx->staticdisttable) {
                    zlib_freetable(dctx, &dctx->currdisttable);
                    dctx->currdisttable = NULL;
                }
            } else if (code < 286) {   /* static tree can give >285; ignore */
                dctx->state = GOTLENSYM;
                dctx->sym = code;
            }
        }
        else if (dctx->state == GOTLENSYM) {
            rec = &rtbl->lencodes[dctx->sym - 257];
            if (dctx->nbits < rec->extrabits)
                goto finished;
            dctx->len =
                rec->min + (dctx->bits & ((1 << rec->extrabits) - 1));
            EATBITS(rec->extrabits);
            dctx->state = GOTLEN;
        }
        else if (dctx->state == GOTLEN) {
            code =
                zlib_huflookup(&dctx->bits, &dctx->nbits,
                               dctx->currdisttable);
            if (code == -1)
                goto finished;
            if (code == -2)
                goto decode_error;
            if (code >= 30)            /* dist symbols 30 and 31 are invalid */
                goto decode_error;
            dctx->state = GOTDISTSYM;
            dctx->sym = code;
        }
        else if (dctx->state == GOTDISTSYM) {
            rec = &rtbl->distcodes[dctx->sym];
            if (dctx->nbits < rec->extrabits)
                goto finished;
            dist = rec->min + (dctx->bits & ((1 << rec->extrabits) - 1));
            EATBITS(rec->extrabits);
            dctx->state = INBLK;
            while (dctx->len--)
                zlib_emit_char(dctx, dctx->window[(dctx->winpos - dist) &
                                                  (WINSIZE - 1)]);
        }
        else if (dctx->state == UNCOMP_LEN) {
            /*
             * Uncompressed block. We expect to see a 16-bit LEN.
             */
            if (dctx->nbits < 16)
                goto finished;
            dctx->uncomplen = dctx->bits & 0xFFFF;
            EATBITS(16);
            dctx->state = UNCOMP_NLEN;
        }
        else if (dctx->state == UNCOMP_NLEN) {
            /*
             * Uncompressed block. We expect to see a 16-bit NLEN,
             * which should be the one's complement of the previous
             * LEN.
             */
            if (dctx->nbits < 16)
                goto finished;
            nlen = dctx->bits & 0xFFFF;
            EATBITS(16);
            if (dctx->uncomplen != (nlen ^ 0xFFFF))
                goto decode_error;
            if (dctx->uncomplen == 0)
                dctx->state = OUTSIDEBLK;       /* block is empty */
            else
                dctx->state = UNCOMP_DATA;
        }
        else if (dctx->state == UNCOMP_DATA) {
            if (dctx->nbits < 8)
                goto finished;
            zlib_emit_char(dctx, dctx->bits & 0xFF);
            EATBITS(8);
            if (--dctx->uncomplen == 0)
                dctx->state = OUTSIDEBLK;       /* end of uncompressed block */
        }
    }

finished:
    *outblock = dctx->outblk;
    *outlen = dctx->outlen;
    return 1;

decode_error:
    sfree(dctx->outblk);
    *outblock = dctx->outblk = NULL;
    *outlen = 0;
    return 0;
}

static struct RelocTable rtable = {
    vbzlib_compress_init,
    vbzlib_compress_cleanup,
    vbzlib_compress_block,
    vbzlib_decompress_init,
    vbzlib_decompress_cleanup,
    vbzlib_decompress_block,
    vbzlib_crc32,
    vbzlib_memnonce,
    vbzlib_memxor,
    0,
    0,
    0,
    zdat_lencodes,
    zdat_distcodes,
    zdat_mirrorbytes,
    zdat_lenlenmap,
    0,
};

struct IoBuffers {
    const unsigned char *block;
    int len;
    unsigned char *outblock;
    int outlen;
    int final;
    int greedy;
    int maxmatch;
    int nicelen;
};

// from libdeflate's crc32_slice4 function
static __inline unsigned int ZLIB_STDCALL
crc32_update_byte(const unsigned int *crc32_table, unsigned int remainder, unsigned char next_byte)
{
    return (remainder >> 8) ^ crc32_table[(unsigned char)remainder ^ next_byte];
}

void ZLIB_STDCALL vbzlib_crc32(struct RelocTable *rtbl, const unsigned char *block, int len, unsigned int *pcrc)
{
    const unsigned int *crc32_table = rtbl->crc32table;
    const unsigned char *p = block;
    const unsigned char *end = block + len;
    const unsigned char *end32;
    unsigned int remainder = *pcrc;

    for (; ((unsigned int)p & 3) && p != end; p++)
        remainder = crc32_update_byte(crc32_table, remainder, *p);

    end32 = p + ((end - p) & ~3);
    for (; p != end32; p += 4) {
        unsigned int v = *((const unsigned int *)p);
        remainder =
            crc32_table[0x300 + (unsigned char)((remainder ^ v) >>  0)] ^
            crc32_table[0x200 + (unsigned char)((remainder ^ v) >>  8)] ^
            crc32_table[0x100 + (unsigned char)((remainder ^ v) >> 16)] ^
            crc32_table[0x000 + (unsigned char)((remainder ^ v) >> 24)];
    }

    for (; p != end; p++)
        remainder = crc32_update_byte(crc32_table, remainder, *p);

    *pcrc = remainder;
}

void ZLIB_STDCALL vbzlib_memnonce(unsigned int *block, unsigned int *nonce, int len, int lParam)
{
    for (; len > 0; len -= 16) {
        if (!++nonce[0])
            ++nonce[1];
        *block++ = nonce[0];
        *block++ = nonce[1];
        *block++ = 0;
        *block++ = 0;
    }
}

void ZLIB_STDCALL vbzlib_memxor(const unsigned char *block, unsigned char *dest, int len, int lParam)
{
    while (len--)
        *dest++ ^= *block++;
}

void *ZLIB_STDCALL vbzlib_compress_init(struct RelocTable *rtbl, int wMsg, int wParam, int lParam)
{
    return zlib_compress_init(rtbl);
}

void ZLIB_STDCALL vbzlib_compress_cleanup(void *handle, int wMsg, int wParam, int lParam)
{
    zlib_compress_cleanup(handle);
}

int ZLIB_STDCALL vbzlib_compress_block(void *handle, struct IoBuffers *buf, unsigned int *pcrc, int lParam)
{
    if (pcrc)
        vbzlib_crc32(((struct LZ77Context *)handle)->rtbl, buf->block, buf->len, pcrc);
    return zlib_compress_block(handle, buf->block, buf->len, &buf->outblock, &buf->outlen, buf->final, buf->greedy, buf->maxmatch, buf->nicelen);
}

void *ZLIB_STDCALL vbzlib_decompress_init(struct RelocTable *rtbl, int wMsg, int wParam, int lParam)
{
    return zlib_decompress_init(rtbl);
}

void ZLIB_STDCALL vbzlib_decompress_cleanup(void *handle, int wMsg, int wParam, int lParam)
{
    zlib_decompress_cleanup(handle);
}

int ZLIB_STDCALL vbzlib_decompress_block(void *handle, struct IoBuffers *buf, unsigned int *pcrc, int lParam)
{
    int result;
    
    result = zlib_decompress_block(handle, buf->block, buf->len, &buf->outblock, &buf->outlen);            
    if (pcrc && result && buf->outblock)
        vbzlib_crc32(((struct zlib_decompress_ctx *)handle)->rtbl, buf->outblock, buf->outlen, pcrc);    
    return result;
}

struct ThunkInfo {
    void *code_start;
    void *code_end;
    void *data_start;
    void *data_end;
};

int __stdcall _DllMainCRTStartup(int hInst, int fdwReason, void *lpReserved);
    
void ZLIB_STDCALL vbzlib_extract_thunk(struct RelocTable *rtbl, struct ThunkInfo *info)
{
    memcpy(rtbl, &rtable, sizeof(rtable));
#if (_MSC_VER == 1200)
    info->code_start = zlib_compress_init;
    info->code_end = vbzlib_extract_thunk;
    info->data_start = (void *)zdat_mirrorbytes;
    info->data_end = (char *)zdat_lenlenmap + sizeof(zdat_lenlenmap);
#else
    info->code_start = vbzlib_crc32;
    info->code_end = (char *)zlib_outbits + 96;
    info->data_start = (void *)zdat_mirrorbytes;
    info->data_end = (char *)zdat_distcodes + sizeof(zdat_distcodes);
#endif
}

static int __stdcall _DllMainCRTStartup(int hInst, int fdwReason, void *lpReserved) {
    return 1;
}
